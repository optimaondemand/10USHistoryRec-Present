# wipe-course.ps1 -- delete all content objects from a Canvas course so it can be rebuilt clean
# from 03_Build/canvas-deploy/*.json.
#
# WHY THIS EXISTS. Course 94 was built under an older naming convention: 93 pages, 90 assignments,
# 24 quizzes, none of whose titles match what the manifests now specify. deploy-canvas.ps1 finds
# existing objects by title, so deploying over that course would create a second copy of everything
# rather than updating it. The alternative was an old-to-new mapping with legacyTitle fields, as
# 11CivicsUSGovtHonors had to build. That is unnecessary here because the course is unpublished with
# no students and no submissions, and 03_Build is the authored source of truth for everything the
# manifests own. Wiping removes the mapping problem entirely.
#
# WHY IT IS A SEPARATE SCRIPT. This is the one destructive operation in the build. Keeping it out of
# deploy-canvas.ps1 means no mistyped flag on the deploy path can ever trigger it.
#
# SAFETY. Dry run by default. Refuses outright to touch a course that is published, has students, or
# has any submission. Requires -Execute AND -IUnderstandThisDeletesContent together to do anything.
#
# Usage:
#   .\wipe-course.ps1                                              # dry run, shows what would go
#   .\wipe-course.ps1 -Execute -IUnderstandThisDeletesContent      # actually deletes
#
# TAKE A SNAPSHOT FIRST:
#   .\deploy-canvas.ps1 -Snapshot ..\..\03_Build\canvas-deploy\live-snapshot-pre-wipe-<date>.json
# That snapshot carries every page body, assignment description and quiz question, and is the only
# recovery path once this script runs. It contains quiz answer keys, so it stays on disk in
# 03_Build and is never committed to this public repo.

param(
    [Parameter(Mandatory=$false)][switch]$Execute,
    [Parameter(Mandatory=$false)][switch]$IUnderstandThisDeletesContent,
    # Titles listed here are never deleted. A page whose title matches is UNPUBLISHED instead.
    # "Appendix B: Teacher Standards Crosswalk" is teacher-facing and, per CLAUDE.md, deliberately
    # not migrated into canvas-deploy. It is currently student-visible and full of SS.912. codes,
    # which the hard rules forbid, so unpublishing both removes the violation and keeps the content.
    [Parameter(Mandatory=$false)][string[]]$PreserveUnpublished = @("Appendix B: Teacher Standards Crosswalk"),
    [Parameter(Mandatory=$false)][string]$Token,
    [Parameter(Mandatory=$false)][string]$TokenPath = $(
        if ($env:CANVAS_TOKEN_FILE) { $env:CANVAS_TOKEN_FILE }
        else { Join-Path $env:USERPROFILE "Desktop\Canvas July Token.odt" }
    )
)

$ErrorActionPreference = 'Stop'
# See deploy-canvas.ps1: without this, PS 5.1's progress bar makes large reads appear to hang.
$ProgressPreference = 'SilentlyContinue'

$base     = "https://optimaoaoteam.instructure.com/api/v1"
$courseId = 94

# --- Auth (same contract as deploy-canvas.ps1) --------------------------------

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
    } else { Get-Content $Path -Raw }
    $flat = ($text -replace '<[^>]+>', ' ') -replace '\s+', ' '
    if ($flat -match '([0-9]+~[A-Za-z0-9]{20,})') { return $Matches[1] }
    throw "Canvas token pattern not found in $Path"
}

$t = if ($Token) { $Token.Trim() } else { Get-CanvasToken -Path $TokenPath }
$headers = @{ Authorization = "Bearer $t"; "User-Agent" = "USHistory2100310-Build/1.0 (+optimaondemand)" }

function Canvas-Get {
    param($Path)
    $all = @(); $uri = "$base$Path"
    while ($uri) {
        $resp = Invoke-WebRequest -Uri $uri -Method GET -Headers $headers -UseBasicParsing
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
# Returns $true if it deleted, $false if the object was already gone.
#
# A 404 here is normal, not an error. Deleting a quiz whose quiz_type is 'assignment' also removes
# the Assignment object backing it, so by the time the assignment loop runs, every reading quiz and
# exam already vanished from the assignments list this script fetched up front. Treating 404 as
# success is also what makes the wipe safely resumable after an interruption.
function Canvas-Delete {
    param($Path)
    try {
        Invoke-RestMethod -Uri "$base$Path" -Method DELETE -Headers $headers | Out-Null
        return $true
    } catch [System.Net.WebException] {
        $code = [int]$_.Exception.Response.StatusCode
        if ($code -eq 404) { return $false }
        throw
    }
}
function Canvas-Json {
    param($Method, $Path, $Obj)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes(($Obj | ConvertTo-Json -Depth 20))
    Invoke-RestMethod -Uri "$base$Path" -Method $Method -Headers $headers `
        -ContentType "application/json; charset=utf-8" -Body $bytes
}

# --- Guards -------------------------------------------------------------------
# These run before anything is read in bulk and before anything is deleted. They are what stop this
# script from ever being pointed at a course that is actually in use.

$course = Canvas-Get "/courses/$courseId`?include[]=total_students"
Write-Host "Course $courseId : $($course.name)"
Write-Host "  workflow_state : $($course.workflow_state)"
Write-Host "  total_students : $($course.total_students)"

if ($course.workflow_state -ne 'unpublished') {
    throw "REFUSING: course $courseId is '$($course.workflow_state)', not 'unpublished'. This script only ever runs against an unpublished course."
}
if ([int]$course.total_students -gt 0) {
    throw "REFUSING: course $courseId has $($course.total_students) students enrolled."
}

$pages       = @(Canvas-Get "/courses/$courseId/pages?per_page=100")
$assignments = @(Canvas-Get "/courses/$courseId/assignments?per_page=100")
$quizzes     = @(Canvas-Get "/courses/$courseId/quizzes?per_page=100")

# A submission anywhere means real student work exists and the "nothing to preserve" premise is
# false. Check every assignment, not a sample.
Write-Host "  checking $($assignments.Count) assignments for submissions ..."
$withSubs = @()
foreach ($a in $assignments) {
    $subs = @(Canvas-Get "/courses/$courseId/assignments/$($a.id)/submissions?per_page=100")
    $real = @($subs | Where-Object { $_.workflow_state -ne 'unsubmitted' -or $_.submitted_at })
    if ($real.Count -gt 0) { $withSubs += "$($a.name) ($($real.Count))" }
}
if ($withSubs.Count -gt 0) {
    Write-Host "REFUSING: real submissions found on:"
    foreach ($w in $withSubs) { Write-Host "  - $w" }
    throw "Refusing to delete content that has student submissions."
}
Write-Host "  no submissions found"
Write-Host ""

# --- Plan ---------------------------------------------------------------------

$toUnpublish = @($pages | Where-Object { $PreserveUnpublished -contains $_.title })
$toDeletePgs = @($pages | Where-Object { $PreserveUnpublished -notcontains $_.title })

Write-Host "PLAN"
Write-Host "  delete pages       : $($toDeletePgs.Count)"
Write-Host "  delete assignments : $($assignments.Count)"
Write-Host "  delete quizzes     : $($quizzes.Count)"
Write-Host "  unpublish, keep    : $($toUnpublish.Count)"
foreach ($p in $toUnpublish) { Write-Host "      - $($p.title)" }
Write-Host "  modules are left in place (names and ids are correct; the manifests reference them)"
Write-Host ""

if (-not ($Execute -and $IUnderstandThisDeletesContent)) {
    Write-Host "DRY RUN. Nothing was changed."
    Write-Host "To execute: .\wipe-course.ps1 -Execute -IUnderstandThisDeletesContent"
    return
}

# --- Execute ------------------------------------------------------------------
# Order matters. Quizzes and assignments are deleted before pages only so the progress output reads
# in the same order as the plan above; Canvas has no dependency requiring it.

$delQ = 0; $delA = 0; $delP = 0; $gone = 0
foreach ($q in $quizzes) {
    if (Canvas-Delete "/courses/$courseId/quizzes/$($q.id)") { $delQ++ } else { $gone++ }
}
Write-Host "  quizzes deleted     : $delQ"
foreach ($a in $assignments) {
    if (Canvas-Delete "/courses/$courseId/assignments/$($a.id)") { $delA++ } else { $gone++ }
}
Write-Host "  assignments deleted : $delA"
foreach ($p in $toDeletePgs) {
    if (Canvas-Delete "/courses/$courseId/pages/$($p.url)") { $delP++ } else { $gone++ }
}
Write-Host "  pages deleted       : $delP"
Write-Host "  already gone (404, expected for quiz-backed assignments): $gone"
foreach ($p in $toUnpublish) {
    Canvas-Json PUT "/courses/$courseId/pages/$($p.url)" @{ wiki_page = @{ published = $false } } | Out-Null
    Write-Host "  unpublished (kept)  $($p.title)"
}

Write-Host ""
Write-Host "Wipe complete. Re-read the course to confirm before deploying."
