# deploy-canvas.ps1 -- HS U.S. History: Reconstruction to Present (CPALMS 2100310), Canvas course 94.
#
# Stage 29 (I2). Ported from 06MJWorldHistory's data-driven driver (which matches this course's
# canvas-deploy/*.json schema exactly: content lives in the JSON, never in this script) and given
# 11CivicsUSGovtHonors' reconciliation machinery (paginated GETs, per-type caches, Upsert-* by
# title, Ensure-ModuleItem) to fix the sibling's idempotency defect: every call in 06's version is
# a bare POST, so a second run duplicates every object in the course.
#
# NOTHING in this file is course content. Every title, body, question, and rubric comes from
# canvas-deploy/*.json, which is itself generated from or authored alongside 03_Build's markdown.
#
# Usage:
#   .\deploy-canvas.ps1 -Directory .            -DryRun            # plan the whole course, no HTTP
#   .\deploy-canvas.ps1 -JsonPath .\week-01.json -DryRun           # plan one module, no HTTP
#   .\deploy-canvas.ps1 -Directory .            -ResolveModuleIds  # read-only GET; patch moduleId
#   .\deploy-canvas.ps1 -JsonPath .\week-01.json                   # live deploy (creates/updates)
#
# Run the weeks in order (week-01 .. week-18, then the exams and capstones) so module item
# positions accumulate predictably. Re-running is safe: existing objects are updated in place.
#
# DO NOT publish course 94 to students until Stage 31 (I3) re-points the iframe URLs broken by
# Stage 3 (B0)'s repo restructure. Individual items are created published; the course is not.

param(
    [Parameter(Mandatory=$false)][string]$JsonPath,
    [Parameter(Mandatory=$false)][string]$Directory,
    [Parameter(Mandatory=$false)][switch]$DryRun,
    [Parameter(Mandatory=$false)][switch]$ResolveModuleIds,
    [Parameter(Mandatory=$false)][string]$Snapshot,
    [Parameter(Mandatory=$false)][string]$Token,
    # Resolution order: -Token, then $env:CANVAS_TOKEN_FILE, then the default below. No absolute
    # personal path is hardcoded, because this file is committed to a public repo.
    [Parameter(Mandatory=$false)][string]$TokenPath = $(
        if ($env:CANVAS_TOKEN_FILE) { $env:CANVAS_TOKEN_FILE }
        else { Join-Path $env:USERPROFILE "Desktop\Canvas July Token.odt" }
    )
)

$ErrorActionPreference = 'Stop'

# Windows PowerShell 5.1 renders a progress bar for every Invoke-WebRequest/Invoke-RestMethod, and
# on large responses that rendering dominates the call: the /pages and /assignments reads for this
# course (93 and 90 objects) hung past two minutes with it on, and returned in under a second with
# it off. This one line is the difference between the snapshot working and appearing to freeze.
$ProgressPreference = 'SilentlyContinue'

$base      = "https://optimaoaoteam.instructure.com/api/v1"
$courseId  = 94
$pagesBase = "https://optimaondemand.github.io/10USHistoryRec-Present"

# Tier group NAMES, never live ids. Ensure-TierGroup resolves name -> id at run time and creates
# the group only if it is absent, so this script carries no course-specific integer that would go
# stale. Names match the four tiers in optima-canvas-assignments.
$TIER_NAMES = @{
    tier1 = "Tier 1: Student & Material"
    tier2 = "Tier 2: Student & Community"
    tier3 = "Tier 3: Student & Teacher"
    tier4 = "Tier 4: Ungraded Scaffolding"
}

# optima-m365-skills Part 7: this page must be linked in EVERY module, not just the front matter,
# so a student enrolling mid-year can find it from wherever they land. The page itself is
# m365-guide.json's own standalone deploy; this constant only names it for module-item wiring.
$M365_PAGE_TITLE = "M365 Setup & File Organization Guide"

# --- Auth --------------------------------------------------------------------
# Token is read at run time from a file OUTSIDE the OneDrive tree, or passed via -Token. It is
# never written into this script, into canvas-deploy/*.json, or into any committed file.
#
# Gotcha already paid for: a token file may put the value on the NEXT line after its label, not
# the same line. A same-line-only parse silently yields an empty token and a 401. Both readers
# below therefore search the whole document for the token PATTERN rather than parsing by line.

function Get-CanvasToken {
    param($Path)
    if (-not (Test-Path $Path)) { throw "Token file not found: $Path (pass -Token instead)" }
    $text = if ([System.IO.Path]::GetExtension($Path) -ieq '.odt') {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $entry = $zip.Entries | Where-Object { $_.FullName -eq 'content.xml' }
            if (-not $entry) { throw "content.xml not found inside $Path" }
            $stream = $entry.Open()
            try { (New-Object System.IO.StreamReader($stream)).ReadToEnd() } finally { $stream.Dispose() }
        } finally { $zip.Dispose() }
    } else {
        Get-Content $Path -Raw
    }
    # Strip markup (.odt) and collapse newlines so a label/value split across lines still matches.
    $flat = ($text -replace '<[^>]+>', ' ') -replace '\s+', ' '
    if ($flat -match '([0-9]+~[A-Za-z0-9]{20,})') { return $Matches[1] }
    throw "Canvas token pattern [0-9]+~[A-Za-z0-9]+ not found in $Path"
}

$script:Headers = $null
function Get-Headers {
    if ($script:Headers) { return $script:Headers }
    $t = if ($Token) { $Token.Trim() } else { Get-CanvasToken -Path $TokenPath }
    # Gotcha already paid for: Canvas 403s with "have not provided a valid user agent" unless an
    # explicit User-Agent is sent. Invoke-RestMethod does not set a usable one by default.
    $script:Headers = @{
        Authorization = "Bearer $t"
        "User-Agent"  = "USHistory2100310-Build/1.0 (+optimaondemand)"
    }
    return $script:Headers
}

# --- Low-level HTTP ----------------------------------------------------------
# Canvas-Json: UTF-8 bytes. Required for pages/quizzes/assignments so the course's accented and
# non-English terms (Mezzogiorno, contadini, padrone, Arbeiter-Zeitung, Forverts) survive.
# Canvas-Form: form-urlencoded. Required for rubrics, discussions, and module items.
# Using the wrong one for a given endpoint is the documented time-waster.

function Canvas-Json {
    param($Method, $Path, $Obj)
    $json  = $Obj | ConvertTo-Json -Depth 30
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    Invoke-RestMethod -Uri "$base$Path" -Method $Method -Headers (Get-Headers) `
        -ContentType "application/json; charset=utf-8" -Body $bytes
}

function Canvas-Form {
    param($Method, $Path, $Hash)
    $pairs = foreach ($k in $Hash.Keys) {
        "$([uri]::EscapeDataString($k))=$([uri]::EscapeDataString([string]$Hash[$k]))"
    }
    Invoke-RestMethod -Uri "$base$Path" -Method $Method -Headers (Get-Headers) `
        -ContentType "application/x-www-form-urlencoded" -Body ($pairs -join '&')
}

# Paginated GET following Canvas's Link header. Used for every reconciliation lookup; a
# single-page GET would miss objects past 100 and re-create them as duplicates.
function Canvas-Get {
    param($Path)
    $all = @()
    $uri = "$base$Path"
    while ($uri) {
        $resp = Invoke-WebRequest -Uri $uri -Method GET -Headers (Get-Headers) -UseBasicParsing
        $all += ($resp.Content | ConvertFrom-Json)
        $next = $null
        if ($resp.Headers['Link']) {
            foreach ($l in ($resp.Headers['Link'] -split ',')) {
                if ($l -match '<([^>]+)>;\s*rel="next"') { $next = $Matches[1] }
            }
        }
        $uri = $next
    }
    return $all
}

# --- Reconciliation caches (lazy, once per run) -------------------------------

$script:PagesCache = $null
$script:QuizzesCache = $null
$script:AssignmentsCache = $null
$script:AssignmentGroupsCache = $null
$script:ModulesCache = $null
$script:ModuleItemsCache = @{}

function Get-PagesCache       { if ($null -eq $script:PagesCache)       { $script:PagesCache       = @(Canvas-Get "/courses/$courseId/pages?per_page=100") }; return $script:PagesCache }
function Get-QuizzesCache     { if ($null -eq $script:QuizzesCache)     { $script:QuizzesCache     = @(Canvas-Get "/courses/$courseId/quizzes?per_page=100") }; return $script:QuizzesCache }
function Get-AssignmentsCache { if ($null -eq $script:AssignmentsCache) { $script:AssignmentsCache = @(Canvas-Get "/courses/$courseId/assignments?per_page=100") }; return $script:AssignmentsCache }
function Get-ModulesCache     { if ($null -eq $script:ModulesCache)     { $script:ModulesCache     = @(Canvas-Get "/courses/$courseId/modules?per_page=100") }; return $script:ModulesCache }
function Get-AssignmentGroupsCache { if ($null -eq $script:AssignmentGroupsCache) { $script:AssignmentGroupsCache = @(Canvas-Get "/courses/$courseId/assignment_groups?per_page=100") }; return $script:AssignmentGroupsCache }

function Get-ModuleItemsCache {
    param($ModuleId)
    if (-not $script:ModuleItemsCache.ContainsKey($ModuleId)) {
        $script:ModuleItemsCache[$ModuleId] = @(Canvas-Get "/courses/$courseId/modules/$ModuleId/items?per_page=100")
    }
    return $script:ModuleItemsCache[$ModuleId]
}

function Ensure-TierGroup {
    param($TierKey)
    $name = $TIER_NAMES[$TierKey]
    if (-not $name) { throw "Unknown tier key '$TierKey'" }
    $existing = (Get-AssignmentGroupsCache) | Where-Object { $_.name -eq $name } | Select-Object -First 1
    if ($existing) { return $existing.id }
    $created = Canvas-Json POST "/courses/$courseId/assignment_groups" @{ name = $name }
    $script:AssignmentGroupsCache += $created
    Write-Host "  Assignment group created: $name -> id $($created.id)"
    return $created.id
}

# --- Tier derivation ----------------------------------------------------------
# The 26 JSONs carry no tier field: Stage 28 (I1) authored them without one, and the four exam
# files are generated output that must not be hand-edited. Tier is therefore derived here, from
# the same signals CLAUDE.md's "Tier classification" section uses, and logged for every item so a
# run is auditable against that section.
#
#   Reading Quiz (quiz, no description)   -> Tier 1  auto-graded, Canvas-native
#   Exam (quiz WITH a description)        -> Tier 3  per optima-hs-social-studies II: place the
#                                                    quarterly exams in the existing Tier 3
#                                                    substantive-assessment group, never a new
#                                                    top-level category. Weight is an OPEN FLAG
#                                                    for the course lead; this only sets the home.
#   Assignment, points > 0                -> Tier 3  cq, pse, principalVr, capstones, Appendix
#                                                    A1/A3/A4/A5. On the eight VR-principal weeks
#                                                    (2,3,7,9,12,15,16,18) principalVr carries 10
#                                                    points and lands here, exactly as CLAUDE.md
#                                                    specifies.
#   Assignment, points = 0                -> Tier 4  supplemental VR (vr1/vr2) and Appendix A2,
#                                                    teacher's discretion, ungraded.
#   Wrapper survey                        -> Tier 4  module intro/outro scaffolding.

function Get-QuizTier   { param($Q) if ($Q.description) { 'tier3' } else { 'tier1' } }
function Get-AssignTier { param($A) if ([double]$A.points -gt 0) { 'tier3' } else { 'tier4' } }

# --- Upsert primitives --------------------------------------------------------
# Each finds the live object by title and PUTs it, or POSTs a new one. This is the fix for the
# sibling's bare-POST defect. Per-type title fields differ and are a documented gotcha:
# Pages wiki_page[title], Quizzes quiz[title], Assignments assignment[NAME] (not title),
# module items form-encoded module_item[title].

function Upsert-Page {
    param($Title, $BodyHtml)
    $existing = (Get-PagesCache) | Where-Object { $_.title -eq $Title } | Select-Object -First 1
    $obj = @{ wiki_page = @{ title = $Title; body = $BodyHtml; published = $true } }
    if ($existing) {
        $resp = Canvas-Json PUT "/courses/$courseId/pages/$($existing.url)" $obj
        Write-Host "  Page updated: $Title"
    } else {
        $resp = Canvas-Json POST "/courses/$courseId/pages" $obj
        $script:PagesCache += $resp
        Write-Host "  Page created: $Title"
    }
    return @{ type = 'Page'; url = $resp.url; id = $resp.page_id }
}

# Replaces a quiz's question list positionally: PUT over the questions that already exist, POST
# the surplus, DELETE any left over from a previous longer version. Re-running never accumulates
# duplicate questions.
function Sync-QuizQuestions {
    param($QuizId, $Title, $Questions, $ExistingQuestions, $PointsEach)
    for ($i = 0; $i -lt $Questions.Count; $i++) {
        $q = $Questions[$i]
        $qBody = @{
            question_name   = $Title
            question_text   = $q.text
            question_type   = $(if ($q.type) { $q.type } else { 'multiple_choice_question' })
            points_possible = $PointsEach
        }
        if ($q.answers) {
            $qBody.answers = @(foreach ($a in $q.answers) {
                if ($a -is [string]) { @{ answer_text = $a; answer_weight = 0 } }
                else { @{ answer_text = $a.text; answer_weight = $(if ($a.correct) { 100 } else { 0 }) } }
            })
        }
        if ($q.feedback) { $qBody.neutral_comments_html = $q.feedback }
        if ($i -lt $ExistingQuestions.Count) {
            Canvas-Json PUT "/courses/$courseId/quizzes/$QuizId/questions/$($ExistingQuestions[$i].id)" @{ question = $qBody } | Out-Null
        } else {
            Canvas-Json POST "/courses/$courseId/quizzes/$QuizId/questions" @{ question = $qBody } | Out-Null
        }
    }
    for ($j = $Questions.Count; $j -lt $ExistingQuestions.Count; $j++) {
        Canvas-Json DELETE "/courses/$courseId/quizzes/$QuizId/questions/$($ExistingQuestions[$j].id)" @{} | Out-Null
        Write-Host "    Deleted surplus question $($ExistingQuestions[$j].id)"
    }
}

function Upsert-Quiz {
    param($Title, $TierGroupId, $Questions, $Description)
    $existing = (Get-QuizzesCache) | Where-Object { $_.title -eq $Title } | Select-Object -First 1
    $quizObj = @{
        title               = $Title
        quiz_type           = 'assignment'
        assignment_group_id = $TierGroupId
        published           = $true
        allowed_attempts    = 1
        scoring_policy      = 'keep_highest'
        show_correct_answers = $true
    }
    if ($Description) { $quizObj.description = $Description }
    if ($existing) {
        $quiz = Canvas-Json PUT "/courses/$courseId/quizzes/$($existing.id)" @{ quiz = $quizObj }
        $existingQuestions = @(Canvas-Get "/courses/$courseId/quizzes/$($quiz.id)/questions?per_page=100")
        Write-Host "  Quiz updated: $Title -> id $($quiz.id) ($($Questions.Count) questions)"
    } else {
        $quiz = Canvas-Json POST "/courses/$courseId/quizzes" @{ quiz = $quizObj }
        $script:QuizzesCache += $quiz
        $existingQuestions = @()
        Write-Host "  Quiz created: $Title -> id $($quiz.id) ($($Questions.Count) questions)"
    }
    Sync-QuizQuestions -QuizId $quiz.id -Title $Title -Questions $Questions `
        -ExistingQuestions $existingQuestions -PointsEach 1
    # Gotcha already paid for: the Quizzes API reporting points_possible 0.0 / question_count 0
    # immediately after this is a caching quirk, not a failure. Re-read later to see real values.
    return @{ type = 'Quiz'; id = $quiz.id }
}

function Upsert-QuizSurvey {
    param($Title, $TierGroupId, $Questions)
    $existing = (Get-QuizzesCache) | Where-Object { $_.title -eq $Title } | Select-Object -First 1
    $quizObj = @{ title = $Title; quiz_type = 'survey'; assignment_group_id = $TierGroupId; published = $true }
    if ($existing) {
        $quiz = Canvas-Json PUT "/courses/$courseId/quizzes/$($existing.id)" @{ quiz = $quizObj }
        $existingQuestions = @(Canvas-Get "/courses/$courseId/quizzes/$($quiz.id)/questions?per_page=100")
        Write-Host "  Wrapper survey updated: $Title -> id $($quiz.id) ($($Questions.Count) questions)"
    } else {
        $quiz = Canvas-Json POST "/courses/$courseId/quizzes" @{ quiz = $quizObj }
        $script:QuizzesCache += $quiz
        $existingQuestions = @()
        Write-Host "  Wrapper survey created: $Title -> id $($quiz.id) ($($Questions.Count) questions)"
    }
    Sync-QuizQuestions -QuizId $quiz.id -Title $Title -Questions $Questions `
        -ExistingQuestions $existingQuestions -PointsEach 0
    return @{ type = 'Quiz'; id = $quiz.id }
}

function Upsert-Assignment {
    param($Title, $Description, $TierGroupId, $Points,
          $SubmissionTypes = @("online_text_entry", "online_upload"))
    # Assignments key on NAME, not title. Matching on .title here would silently miss every
    # existing assignment and re-create the lot.
    $existing = (Get-AssignmentsCache) | Where-Object { $_.name -eq $Title } | Select-Object -First 1
    $assignObj = @{
        name                = $Title
        description         = $Description
        points_possible     = $Points
        submission_types    = $SubmissionTypes
        assignment_group_id = $TierGroupId
        published           = $true
    }
    if ($existing) {
        $resp = Canvas-Json PUT "/courses/$courseId/assignments/$($existing.id)" @{ assignment = $assignObj }
        Write-Host "  Assignment updated: $Title -> id $($resp.id) ($Points pts)"
    } else {
        $resp = Canvas-Json POST "/courses/$courseId/assignments" @{ assignment = $assignObj }
        $script:AssignmentsCache += $resp
        Write-Host "  Assignment created: $Title -> id $($resp.id) ($Points pts)"
    }
    return @{ type = 'Assignment'; id = $resp.id }
}

# --- Rubrics ------------------------------------------------------------------
# optima-canvas-assignments requires every Tier 3 item to be graded with a native Canvas rubric of
# 3-4 criteria, held consistent across modules so grading stays near five minutes per student.
#
# Two sources, decided with the course lead 2026-08-01:
#   1. If the JSON assignment carries its own `rubric` object, use it verbatim. Only the two
#      quarter capstones do (4 criteria x 25 = 100, with long_description on each).
#   2. Otherwise build the standard 4-criterion rubric and scale it to the assignment's own
#      points, so a 10-point Critical Questions gets 2.5 per criterion and the totals reconcile.
# Criterion names are optima-canvas-assignments' four, matching both sibling courses.

$STANDARD_CRITERIA = @(
    "Understanding of the material",
    "Quality of reasoning",
    "Clarity and precision of expression",
    "Intellectual honesty (naming the limits of your own understanding)"
)

function Build-RubricForm {
    param($Title, $AssignmentId, $Criteria)
    $form = [ordered]@{
        "rubric[title]"                          = "$Title Rubric"
        "rubric[free_form_criterion_comments]"   = "true"
        "rubric_association[association_id]"     = $AssignmentId
        "rubric_association[association_type]"   = "Assignment"
        "rubric_association[use_for_grading]"    = "true"
        "rubric_association[purpose]"            = "grading"
    }
    for ($i = 0; $i -lt $Criteria.Count; $i++) {
        $c   = $Criteria[$i]
        $pts = [double]$c.points
        $form["rubric[criteria][$i][description]"] = $c.description
        $form["rubric[criteria][$i][points]"]      = $pts
        if ($c.long_description) { $form["rubric[criteria][$i][long_description]"] = $c.long_description }
        $form["rubric[criteria][$i][ratings][0][description]"] = "Full"
        $form["rubric[criteria][$i][ratings][0][points]"]      = $pts
        $form["rubric[criteria][$i][ratings][1][description]"] = "Partial"
        $form["rubric[criteria][$i][ratings][1][points]"]      = [math]::Round($pts * 0.6, 2)
        $form["rubric[criteria][$i][ratings][2][description]"] = "Minimal"
        $form["rubric[criteria][$i][ratings][2][points]"]      = 0
    }
    return $form
}

# Splits $Points across the four standard criteria without rounding drift: the last criterion
# absorbs any remainder so the rubric total always equals the assignment's points exactly.
function Get-StandardCriteria {
    param($Points)
    $each = [math]::Round(([double]$Points) / 4.0, 2)
    $out = @()
    for ($i = 0; $i -lt 4; $i++) {
        $p = if ($i -eq 3) { [math]::Round(([double]$Points) - ($each * 3), 2) } else { $each }
        $out += @{ description = $STANDARD_CRITERIA[$i]; points = $p }
    }
    return $out
}

function Ensure-Rubric {
    param($AssignmentId, $Title, $JsonRubric, $Points)
    $live = Canvas-Get "/courses/$courseId/assignments/$AssignmentId"
    if ($live.rubric) { Write-Host "    Rubric already attached: $Title"; return }
    if ($JsonRubric) {
        $criteria = $JsonRubric.criteria
        $label    = "$($criteria.Count)-criterion rubric from JSON"
    } else {
        $criteria = Get-StandardCriteria -Points $Points
        $label    = "standard 4-criterion rubric scaled to $Points pts"
    }
    $form = Build-RubricForm -Title $Title -AssignmentId $AssignmentId -Criteria $criteria
    Canvas-Form POST "/courses/$courseId/rubrics" $form | Out-Null
    Write-Host "    Rubric attached ($label): $Title"
}

# --- Iframe wrapper -----------------------------------------------------------
# GitHub-hosted lesson pages are EMBEDDED, never merely linked. Linking instead of embedding is
# the exact failure mode optima-hs-social-studies names as never to repeat.

function Iframe-Wrapper {
    param($Title, $Url, $Height)
    # No fallback-footer link to the raw GitHub Pages URL here, by design (2026-08-11).
    # A prior version of this function generated one ("If the page doesn't load, open it in a
    # new tab") and every page built by it carried a live, working second route straight out of
    # Canvas to GitHub Pages, bypassing the iframe entirely -- which is exactly the pattern
    # teachers flagged as alarming when a student clicked it (see the intra-course link audit
    # skill and this course's own remediation history). All 88 existing pages were fixed by
    # stripping this div directly via the Canvas API; that fix is NOT durable against a future
    # deploy unless this function stays footer-free, since every live deploy regenerates the page
    # body from this template. Do not re-add a footer link here.
    return "<div style=`"margin: 20px 0; border: 2px solid #2e86c1; border-radius: 12px; overflow: hidden; background: #f8f9fa;`">`n  <div style=`"background: #1a5276; color: #fff; padding: 10px 18px; font-family: Arial, Helvetica, sans-serif; font-weight: bold; font-size: 1.05em;`">$Title</div>`n  <iframe src=`"$Url`" width=`"100%`" height=`"$Height`" style=`"border: none; display: block;`" allowfullscreen></iframe>`n</div>"
}

# --- Module item wiring -------------------------------------------------------
# Never re-wires an item already linked in the module, and appends new items after the current
# highest position rather than trusting a per-run counter, so this is safe to re-run and safe
# across several JSON files that share one module.
#
# The item title is set explicitly. Canvas defaults a module item's title to its content object's
# title at creation only: it does NOT follow a later rename of that object. Setting it here, and
# correcting it when it has drifted, keeps the module list matching the twelve-slot numbering.

function Ensure-ModuleItem {
    param($ModuleId, $Key, $Item, $Title, $Position)
    $existingItems = Get-ModuleItemsCache -ModuleId $ModuleId
    $already = $existingItems | Where-Object {
        ($Item.type -eq 'Page' -and $_.type -eq 'Page' -and $_.page_url -eq $Item.url) -or
        ($Item.type -ne 'Page' -and $_.type -eq $Item.type -and $_.content_id -eq $Item.id)
    } | Select-Object -First 1

    if ($already) {
        if ($already.title -ne $Title) {
            Canvas-Form PUT "/courses/$courseId/modules/$ModuleId/items/$($already.id)" `
                @{ "module_item[title]" = $Title } | Out-Null
            Write-Host "  Module item retitled: '$($already.title)' -> '$Title'"
        } else {
            Write-Host "  Module item already wired: $Key ($($Item.type))"
        }
        return
    }

    $nextPosition = if ($Position) { $Position }
                    elseif ($existingItems.Count -gt 0) { (($existingItems | Measure-Object -Property position -Maximum).Maximum) + 1 }
                    else { 1 }
    $form = [ordered]@{
        "module_item[position]" = $nextPosition
        "module_item[title]"    = $Title
    }
    if ($Item.type -eq 'Page') {
        $form["module_item[type]"]     = "Page"
        $form["module_item[page_url]"] = $Item.url
    } else {
        $form["module_item[type]"]       = $Item.type
        $form["module_item[content_id]"] = $Item.id
    }
    $resp = Canvas-Form POST "/courses/$courseId/modules/$ModuleId/items" $form
    $script:ModuleItemsCache[$ModuleId] += $resp
    Write-Host "  Module item added at position $nextPosition : $Key ($($Item.type))"
}

# --- M365 module wiring --------------------------------------------------------
# optima-m365-skills Part 7: the M365 guide must be linked in every module, not just the front
# matter (module 1304, where it already sits at 0.5). Weekly modules keep optima-hs-social-studies
# V.3's module-order naming ("N.13", the 13th slot after the twelve-slot grid); the capstone, exam,
# and appendix modules are not number-prefixed at all, so they get the plain page title.

function Get-M365ItemTitle {
    param($Week)
    $isWeekly = ($Week -is [int]) -or ("$Week" -match '^\d+$')
    if ($isWeekly) { return "$Week.13 $M365_PAGE_TITLE" }
    return $M365_PAGE_TITLE
}

# Resolves the live M365 page by title (never created here: it is m365-guide.json's own deploy).
# Fails loudly rather than skipping, because a module silently missing this page is itself a
# compliance failure, not a condition to route around.
function Resolve-M365PageItem {
    $page = (Get-PagesCache) | Where-Object { $_.title -eq $M365_PAGE_TITLE } | Select-Object -First 1
    if (-not $page) {
        throw "M365 page '$M365_PAGE_TITLE' not found in course $courseId. Deploy m365-guide.json before wiring module items, per optima-m365-skills Part 7."
    }
    return @{ type = 'Page'; url = $page.url; id = $page.page_id }
}

# Appends the M365 page as the LAST item of the module (never position 1, which the intro wrapper
# owns). Ensure-ModuleItem is reused verbatim, so this is idempotent for free: a module that
# already carries the page (1304) is left untouched beyond a title check.
function Ensure-M365ModuleItem {
    param($ModuleId, $Week)
    $item  = Resolve-M365PageItem
    $title = Get-M365ItemTitle -Week $Week
    Ensure-ModuleItem -ModuleId $ModuleId -Key 'm365' -Item $item -Title $title
}

# --- Manifest loading ---------------------------------------------------------

function Get-ManifestFiles {
    if ($JsonPath) { return @((Resolve-Path $JsonPath).Path) }
    if (-not $Directory) { throw "Pass -JsonPath <file> or -Directory <dir>." }
    # Weeks first and in order, then the standalone modules, so positions accumulate predictably.
    #
    # Only files that actually look like manifests are returned. A snapshot or any other stray .json
    # dropped in this directory would otherwise be treated as a module and crash the run: that is
    # exactly what happened when live-snapshot-pre-wipe-2026-08-01.json was first written here.
    # Snapshots now live in 03_Build/canvas-snapshots/, and this guard keeps the mistake from
    # mattering if one lands here again.
    # The test is a raw text match, not ConvertFrom-Json. PS 5.1's JSON parser is slow enough on
    # these files (the semester final is 44 KB and deeply nested) that parsing every candidate here
    # added minutes to a dry run that otherwise takes seconds.
    $dir = (Resolve-Path $Directory).Path
    $all = @(Get-ChildItem -Path $dir -Filter '*.json')
    $manifests = @($all | Where-Object {
        $head = Get-Content $_.FullName -Raw -Encoding UTF8
        $head -match '"moduleOrder"\s*:' -or ($head -match '"title"\s*:' -and $head -match '"body"\s*:')
    })
    $weeks = @($manifests | Where-Object { $_.Name -like 'week-*' } | Sort-Object Name)
    $rest  = @($manifests | Where-Object { $_.Name -notlike 'week-*' } | Sort-Object Name)
    $skipped = $all.Count - $manifests.Count
    if ($skipped -gt 0) { Write-Host "Skipped $skipped non-manifest .json file(s) in $dir`n" }
    return @($weeks.FullName) + @($rest.FullName)
}

# --- Snapshot -----------------------------------------------------------------
# Read-only. Captures the complete live state of the course to one JSON file: every module with its
# items, and every page, assignment and quiz WITH ITS FULL BODY and, for quizzes, its questions.
#
# Bodies are the point. A snapshot of titles and ids alone would be useless as a backup, and this
# file is the only recovery path once wipe-course.ps1 runs. The list endpoints do not return page
# bodies, so each page is re-fetched individually by url.

function Invoke-Snapshot {
    param($Path)
    Write-Host "Capturing live state of course $courseId ..."

    $modules = @()
    foreach ($m in (Get-ModulesCache)) {
        $modules += [ordered]@{
            id = $m.id; name = $m.name; position = $m.position
            published = $m.published
            items = @(Canvas-Get "/courses/$courseId/modules/$($m.id)/items?per_page=100")
        }
    }
    Write-Host "  modules: $($modules.Count) (items: $(($modules | ForEach-Object { $_.items.Count } | Measure-Object -Sum).Sum))"

    # /pages omits `body`; each page must be fetched by its url to get one.
    $pages = @()
    foreach ($p in (Get-PagesCache)) {
        $full = Canvas-Get "/courses/$courseId/pages/$($p.url)"
        $pages += [ordered]@{
            page_id = $p.page_id; url = $p.url; title = $p.title
            published = $p.published; body = $full.body
        }
    }
    Write-Host "  pages: $($pages.Count)"

    $assignments = @()
    foreach ($a in (Get-AssignmentsCache)) {
        $assignments += [ordered]@{
            id = $a.id; name = $a.name; published = $a.published
            points_possible = $a.points_possible
            assignment_group_id = $a.assignment_group_id
            submission_types = $a.submission_types
            description = $a.description
            rubric = $a.rubric
        }
    }
    Write-Host "  assignments: $($assignments.Count)"

    $quizzes = @()
    foreach ($q in (Get-QuizzesCache)) {
        $quizzes += [ordered]@{
            id = $q.id; title = $q.title; quiz_type = $q.quiz_type
            published = $q.published; description = $q.description
            assignment_group_id = $q.assignment_group_id
            questions = @(Canvas-Get "/courses/$courseId/quizzes/$($q.id)/questions?per_page=100")
        }
    }
    Write-Host "  quizzes: $($quizzes.Count)"

    $snap = [ordered]@{
        _note = "Read-only capture of Canvas course $courseId taken before wipe-course.ps1 ran. Bodies, descriptions and quiz questions are included, so this is a genuine backup and not merely an inventory."
        courseId = $courseId
        capturedBy = "deploy-canvas.ps1 -Snapshot"
        assignmentGroups = @(Get-AssignmentGroupsCache | ForEach-Object { [ordered]@{ id = $_.id; name = $_.name; group_weight = $_.group_weight } })
        modules = $modules
        pages = $pages
        assignments = $assignments
        quizzes = $quizzes
    }
    $json = $snap | ConvertTo-Json -Depth 40
    # .NET resolves a relative path against the process working directory, which is NOT PowerShell's
    # current location. Make it absolute first, or the file silently lands somewhere unexpected.
    $abs = [System.IO.Path]::GetFullPath((Join-Path (Get-Location).ProviderPath $Path))
    [System.IO.File]::WriteAllText($abs, $json, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host "Snapshot written: $abs ($([math]::Round((Get-Item $abs).Length / 1KB)) KB)"
}

# --- ResolveModuleIds ---------------------------------------------------------
# Read-only. GETs course 94's module list and matches each manifest to a live module by name,
# then patches moduleId into the manifest file on disk. No Canvas object is created or modified,
# so this honors Stage 29's "no Canvas writes" rule while unblocking Stage 31.

function Invoke-ResolveModuleIds {
    param($Files)
    $modules = Get-ModulesCache
    Write-Host "Course $courseId has $($modules.Count) live modules:"
    foreach ($m in ($modules | Sort-Object position)) {
        Write-Host ("  [{0,3}] id={1,-6} {2}" -f $m.position, $m.id, $m.name)
    }
    Write-Host ""
    $resolved = 0; $unresolved = @()
    foreach ($f in $Files) {
        $name = [System.IO.Path]::GetFileName($f)
        $raw  = Get-Content $f -Raw -Encoding UTF8
        $data = $raw | ConvertFrom-Json
        if ($null -eq $data.PSObject.Properties['moduleId']) { Write-Host "  $name : no moduleId field, skipped"; continue }
        if ($null -ne $data.moduleId) { Write-Host "  $name : already set to $($data.moduleId)"; continue }

        # Weekly manifests match a module whose name contains "Module N" or "Week N" for their own
        # number. Standalone manifests are matched through the explicit alias map below rather than
        # by fuzzy name search, so every pairing is auditable.
        #
        # All four exams share one module, "Exams and Final", created 2026-08-01 during the wipe and
        # rebuild. Course 94 previously had no exam module at all, which is why these four manifests
        # carried moduleId null through Stage 29.
        $ALIASES = @{
            'q1-capstone.json'       = '(?i)^Q1 Capstone\b'
            'q2-capstone.json'       = '(?i)^Q2 Capstone\b'
            'appendices.json'        = '(?i)^Appendix\b'
            'q1-exam.json'           = '(?i)^Exams and Final$'
            'q2-exam.json'           = '(?i)^Exams and Final$'
            'semester-final.json'    = '(?i)^Exams and Final$'
            'eoc-practice-exam.json' = '(?i)^Exams and Final$'
        }
        $match = $null
        if ($ALIASES.ContainsKey($name)) {
            $match = $modules | Where-Object { $_.name -match $ALIASES[$name] } | Select-Object -First 1
        } elseif ($data.week -is [int] -or $data.week -match '^\d+$') {
            $n = [int]$data.week
            $match = $modules | Where-Object { $_.name -match "(?i)\b(module|week)\s*0*$n\b" } | Select-Object -First 1
        }
        if ($match) {
            $patched = $raw -replace '("moduleId"\s*:\s*)null', "`${1}$($match.id)"
            if ($patched -eq $raw) { throw "Could not patch moduleId in $name" }
            # Retire the stale authoring-time TODO rather than leaving it asserting that the id
            # is still unresolved. Both keys are underscore-prefixed build metadata, exempt from
            # tools/verify-canvas-deploy.js and never student-facing.
            $noteEsc = [regex]::Escape($match.name) -replace '"', '\"'
            $patched = $patched -replace '"_moduleId_TODO"\s*:\s*"[^"]*"',
                "`"_moduleId_resolved`": `"Stage 29 (I2), 2026-08-01: matched to live course 94 module $($match.id) by name, via deploy-canvas.ps1 -ResolveModuleIds (read-only GET; no Canvas object was created or changed).`""
            # WriteAllText with an explicit BOM-less encoder, NOT Set-Content -Encoding UTF8.
            # Windows PowerShell 5.1's utf8 writes a BOM, and a BOM makes the file unparseable
            # to JSON.parse, which is what tools/verify-canvas-deploy.js and the exam converter
            # both use. This corrupted all 21 patched manifests the first time it ran.
            [System.IO.File]::WriteAllText($f, $patched, (New-Object System.Text.UTF8Encoding($false)))
            Write-Host "  $name : moduleId -> $($match.id)  ($($match.name))"
            $resolved++
        } else {
            $unresolved += $name
        }
    }
    Write-Host ""
    Write-Host "Resolved $resolved manifest(s)."
    if ($unresolved.Count -gt 0) {
        Write-Host "UNRESOLVED (need a module chosen by hand before deploying):"
        foreach ($u in $unresolved) { Write-Host "  - $u" }
    }
}

# --- Deploy one manifest -------------------------------------------------------

function Invoke-Manifest {
    param($File)
    $data = Get-Content $File -Raw -Encoding UTF8 | ConvertFrom-Json
    $name = [System.IO.Path]::GetFileName($File)

    # m365-guide.json is a standalone Canvas-native Page (title + body, no module, no moduleOrder).
    # Its body is already RCE-safe inline HTML and must NOT be iframe-wrapped.
    if ($data.title -and $data.body -and -not $data.moduleOrder) {
        Write-Host "=== $name : standalone page ==="
        if ($DryRun) { Write-Host "  [dry-run] Page: $($data.title) ($($data.body.Length) chars, Canvas-native)"; return }
        Upsert-Page -Title $data.title -BodyHtml $data.body | Out-Null
        return
    }

    Write-Host "=== $name : week $($data.week) : module $($data.moduleId) : gh '$($data.ghFolder)' ==="
    if (-not $DryRun -and $null -eq $data.moduleId) {
        throw "$name has moduleId null. Run with -ResolveModuleIds first, or set it by hand."
    }

    $created = @{}
    $titles  = @{}

    foreach ($p in $data.pages) {
        $url  = "$pagesBase/$($data.ghFolder)/$($p.file)"
        $body = Iframe-Wrapper -Title $p.title -Url $url -Height $p.height
        $titles[$p.key] = $p.title
        if ($DryRun) { Write-Host "  [dry-run] Page  : $($p.title)  <- $url"; $created[$p.key] = @{ type='Page'; url='(dry-run)' }; continue }
        $created[$p.key] = Upsert-Page -Title $p.title -BodyHtml $body
    }

    foreach ($q in $data.quizzes) {
        $tier = Get-QuizTier -Q $q
        $titles[$q.key] = $q.title
        if ($DryRun) { Write-Host "  [dry-run] Quiz  : $($q.title)  [$tier] $($q.questions.Count) q$(if($q.description){', has description'})"; $created[$q.key] = @{ type='Quiz'; id=0 }; continue }
        $gid = Ensure-TierGroup -TierKey $tier
        $created[$q.key] = Upsert-Quiz -Title $q.title -TierGroupId $gid -Questions $q.questions -Description $q.description
    }

    foreach ($a in $data.assignments) {
        $tier = Get-AssignTier -A $a
        $titles[$a.key] = $a.title
        $rub = if ($a.rubric) { "JSON rubric ($($a.rubric.criteria.Count) criteria)" }
               elseif ($tier -eq 'tier3') { "standard rubric" } else { "no rubric (Tier 4)" }
        if ($DryRun) { Write-Host "  [dry-run] Assign: $($a.title)  [$tier] $($a.points) pts, $rub"; $created[$a.key] = @{ type='Assignment'; id=0 }; continue }
        $gid  = Ensure-TierGroup -TierKey $tier
        $resp = Upsert-Assignment -Title $a.title -Description $a.body -TierGroupId $gid -Points $a.points
        if ($tier -eq 'tier3') { Ensure-Rubric -AssignmentId $resp.id -Title $a.title -JsonRubric $a.rubric -Points $a.points }
        $created[$a.key] = $resp
    }

    foreach ($w in $data.wrappers) {
        $titles[$w.key] = $w.title
        if ($DryRun) { Write-Host "  [dry-run] Survey: $($w.title)  [tier4] $($w.questions.Count) q"; $created[$w.key] = @{ type='Quiz'; id=0 }; continue }
        $gid = Ensure-TierGroup -TierKey 'tier4'
        $created[$w.key] = Upsert-QuizSurvey -Title $w.title -TierGroupId $gid -Questions $w.questions
    }

    # moduleOrder integrity: every key must have been built, or the wiring below would silently
    # skip an item and leave a gap in the module.
    foreach ($key in $data.moduleOrder) {
        if (-not $created.ContainsKey($key)) { throw "$name : moduleOrder key '$key' has no matching page/quiz/assignment/wrapper" }
    }

    if ($DryRun) {
        Write-Host "  [dry-run] moduleOrder ($($data.moduleOrder.Count) items): $($data.moduleOrder -join ' ')"
        Write-Host "  [dry-run] M365    : $(Get-M365ItemTitle -Week $data.week)  (appended last, no HTTP)"
        return
    }

    $position = 1
    foreach ($key in $data.moduleOrder) {
        Ensure-ModuleItem -ModuleId $data.moduleId -Key $key -Item $created[$key] -Title $titles[$key] -Position $position
        $position++
    }

    Ensure-M365ModuleItem -ModuleId $data.moduleId -Week $data.week

    Write-Host "=== $name complete (reconciled, no duplicates created) ==="
}

# --- Main ----------------------------------------------------------------------

# -Snapshot needs no manifests, so it is handled before Get-ManifestFiles insists on one.
if ($Snapshot) {
    Invoke-Snapshot -Path $Snapshot
    return
}

$files = Get-ManifestFiles

if ($ResolveModuleIds) {
    Invoke-ResolveModuleIds -Files $files
    return
}

if ($DryRun) { Write-Host "DRY RUN: no HTTP request will be made.`n" }

foreach ($f in $files) {
    Invoke-Manifest -File $f
    Write-Host ""
}

Write-Host "All done ($($files.Count) manifest(s))."
