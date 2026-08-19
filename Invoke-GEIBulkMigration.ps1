			'By using the following materials or sample code you agree to be bound by the license terms below 
			'and the Microsoft Partner Program Agreement the terms of which are incorporated herein by this reference. 
			'These license terms are an agreement between Microsoft Corporation (or, if applicable based on where you 
			'are located, one of its affiliates) and you. Any materials (other than sample code) we provide to you 
			'are for your internal use only. Any sample code is provided for the purpose of illustration only and is 
			'not intended to be used in a production environment. We grant you a nonexclusive, royalty-free right to 
			'use and modify the sample code and to reproduce and distribute the object code form of the sample code, 
			'provided that you agree: (i) to not use Microsoft’s name, logo, or trademarks to market your software product 
			'in which the sample code is embedded; (ii) to include a valid copyright notice on your software product in 
			'which the sample code is embedded; (iii) to provide on behalf of and for the benefit of your subcontractors 
			'a disclaimer of warranties, exclusion of liability for indirect and consequential damages and a reasonable 
			'limitation of liability; and (iv) to indemnify, hold harmless, and defend Microsoft, its affiliates and 
			'suppliers from and against any third party claims or lawsuits, including attorneys’ fees, that arise or result 
'from the use or distribution of the sample code."

<#
.SYNOPSIS
    Bulk-migrates repositories from GitHub Enterprise Server or github.com to GitHub Enterprise
    Cloud (github.com).

.DESCRIPTION
    Reads a CSV of repositories, queues a `gh gei migrate-repo` for each one, then waits for
    every queued migration to finish. Results are written to a timestamped CSV report.

    Use -SourceType GHEC to test with source repositories on github.com. Use -SourceType GHES
    (the default) for customer migrations from GitHub Enterprise Server.

    Use this script when the destination is standard GitHub Enterprise Cloud on github.com.
   
    Required environment variables (never pass these as parameters):
      GH_PAT                            - Classic PAT for the TARGET org on github.com.
                                          Scopes: repo, admin:org, workflow (org owner)
                                                  repo, read:org, workflow (migrator role)
          GH_SOURCE_PAT                     - Classic PAT for the SOURCE org on GHES or github.com.
                          Scopes: repo, admin:org
      AZURE_STORAGE_CONNECTION_STRING   - Azure Blob storage used to stage the migration archive.
                                          GEI reads this variable directly, so it is not passed on
                                          the command line where it could leak into logs/history.
                          Only used for GHES sources. Not required when
                          -UseGitHubStorage is specified.

.PARAMETER RepoListPath
    Path to a CSV with columns: SourceOrg, SourceRepo, TargetOrg, TargetRepo.
    TargetOrg and TargetRepo are optional and default to DefaultTargetOrg / SourceRepo.

.PARAMETER SourceType
    Source repository platform. Use GHEC for github.com or GHES for GitHub Enterprise Server.

.PARAMETER GhesApiUrl
    API endpoint of the GitHub Enterprise Server instance. Used only when SourceType is GHES.

.PARAMETER DefaultTargetOrg
    Target org used for rows that leave TargetOrg blank.

.PARAMETER OutputPath
    Directory for the results CSV and per-repo logs.

.PARAMETER ThrottleLimit
    Maximum number of migrations in flight at once. GitHub allows up to 5 concurrent repository
    migrations per organization.

.EXAMPLE
    .\Invoke-GeiBulkMigration.ps1 -RepoListPath .\repos.csv -SourceType GHEC -WhatIf

.EXAMPLE
    .\Invoke-GeiBulkMigration.ps1 -RepoListPath .\repos.csv -SkipExisting -ThrottleLimit 3
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
    [string]$RepoListPath,

    [ValidateSet('GHES', 'GHEC')]
    [string]$SourceType = 'GHES',

    [ValidatePattern('^https://')]
    [string]$GhesApiUrl = '#API URL#',

    [string]$DefaultTargetOrg = '# ORG TARGET#',

    [string]$OutputPath = (Join-Path $PSScriptRoot 'logs'),

    [ValidateRange(1, 5)]
    [int]$ThrottleLimit = 3,

    [ValidateSet('private', 'internal', 'public')]
    [string]$TargetRepoVisibility,

    # Skip repos that already exist in the target org instead of failing.
    [switch]$SkipExisting,

    # Use GitHub-owned blob storage instead of your own Azure account (GEI CLI v1.9.0+).
    [switch]$UseGitHubStorage,

    # For GHES instances with a self-signed or otherwise invalid TLS certificate.
    [switch]$NoSslVerify
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$script:TargetHost = 'github.com'

#region Preflight -------------------------------------------------------------

function Assert-Prerequisites {
    if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
        throw "GitHub CLI ('gh') was not found on PATH. Install it from https://cli.github.com/"
    }

    # Out-String collapses the result to one string; -notmatch on an array returns elements, not a bool.
    $extensions = (gh extension list 2>&1 | Out-String)
    if ($LASTEXITCODE -ne 0 -or $extensions -notmatch 'gei') {
        throw "The 'gei' extension is not installed. Run: gh extension install github/gh-gei"
    }

    $required = @('GH_PAT', 'GH_SOURCE_PAT')
    if ($SourceType -eq 'GHES' -and -not $UseGitHubStorage) {
        $required += 'AZURE_STORAGE_CONNECTION_STRING'
    }

    $missing = $required |
        Where-Object { [string]::IsNullOrWhiteSpace([Environment]::GetEnvironmentVariable($_)) }

    if ($missing) {
        throw "Missing required environment variable(s): $($missing -join ', ')"
    }
}

function Import-RepoList {
    param([string]$Path)

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { throw "'$Path' contains no rows." }

    $columns = $rows[0].PSObject.Properties.Name
    $missingCols = @('SourceOrg', 'SourceRepo') | Where-Object { $_ -notin $columns }
    if ($missingCols) {
        throw "CSV is missing required column(s): $($missingCols -join ', ')"
    }

    $rows | ForEach-Object {
        $sourceOrg = "$($_.SourceOrg)".Trim()
        $sourceRepo = "$($_.SourceRepo)".Trim()
        if (-not $sourceOrg -or -not $sourceRepo) {
            throw 'CSV contains a row with an empty SourceOrg or SourceRepo.'
        }

        $targetOrg = if ($columns -contains 'TargetOrg' -and $_.TargetOrg) { "$($_.TargetOrg)".Trim() } else { $DefaultTargetOrg }
        $targetRepo = if ($columns -contains 'TargetRepo' -and $_.TargetRepo) { "$($_.TargetRepo)".Trim() } else { $sourceRepo }

        [pscustomobject]@{
            SourceOrg   = $sourceOrg
            SourceRepo  = $sourceRepo
            TargetOrg   = $targetOrg
            TargetRepo  = $targetRepo
            MigrationId = ''
            Status      = 'Pending'
            Detail      = ''
            StartedAt   = $null
            FinishedAt  = $null
        }
    }
}

function Get-RepoLogPath {
    param([pscustomobject]$Repo, [string]$Folder)

    $name = "$($Repo.SourceOrg)_$($Repo.SourceRepo)"
    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) { $name = $name.Replace($c, '_') }
    return (Join-Path $Folder "$name.log")
}

function Test-TargetRepoExists {
    param([string]$Org, [string]$Repo)

    # `gh api` authenticates separately from the gei extension; a runner has no `gh auth login`
    # state, so fall back to GH_PAT for this call only and leave the caller's environment alone.
    $previous = $env:GH_TOKEN
    if (-not $previous) { $env:GH_TOKEN = $env:GH_PAT }
    try {
        gh api --hostname $script:TargetHost "repos/$Org/$Repo" --silent 2>$null | Out-Null
        return ($LASTEXITCODE -eq 0)
    }
    finally {
        $env:GH_TOKEN = $previous
    }
}

#endregion

#region Migration -------------------------------------------------------------

function Start-RepoMigration {
    <# Queues one migration and returns the migration id emitted by GEI. #>
    param([pscustomobject]$Repo, [string]$LogFile)

    $arguments = @(
        'gei', 'migrate-repo'
        '--github-source-org', $Repo.SourceOrg
        '--source-repo', $Repo.SourceRepo
        '--github-target-org', $Repo.TargetOrg
        '--target-repo', $Repo.TargetRepo
        '--queue-only'
    )

    if ($SourceType -eq 'GHES') { $arguments += @('--ghes-api-url', $GhesApiUrl) }
    if ($TargetRepoVisibility) { $arguments += @('--target-repo-visibility', $TargetRepoVisibility) }
    if ($SourceType -eq 'GHES' -and $UseGitHubStorage) { $arguments += '--use-github-storage' }
    if ($SourceType -eq 'GHES' -and $NoSslVerify) { $arguments += '--no-ssl-verify' }

    $output = & gh @arguments 2>&1
    $output | Out-File -LiteralPath $LogFile -Encoding utf8

    if ($LASTEXITCODE -ne 0) {
        throw ($output | Select-Object -Last 5 | Out-String).Trim()
    }

    $match = [regex]::Match(($output | Out-String), 'RM_[A-Za-z0-9_\-=]+')
    if (-not $match.Success) {
        throw "Migration was queued but no migration ID could be parsed. See $LogFile"
    }

    return $match.Value
}

function Wait-RepoMigration {
    param([string]$MigrationId, [string]$LogFile)

    $output = & gh gei wait-for-migration --migration-id $MigrationId 2>&1
    $output | Out-File -LiteralPath $LogFile -Encoding utf8 -Append

    if ($LASTEXITCODE -ne 0) {
        return @{ Succeeded = $false; Detail = ($output | Select-Object -Last 5 | Out-String).Trim() }
    }

    return @{ Succeeded = $true; Detail = 'Migration completed.' }
}

function Complete-OneMigration {
    <# Waits on the oldest in-flight migration and records its outcome. #>
    param([object[]]$Repos, [string]$RunFolder)

    $next = @($Repos | Where-Object { $_.Status -eq 'Queued' })[0]
    if (-not $next) { return }

    $result = Wait-RepoMigration -MigrationId $next.MigrationId -LogFile (Get-RepoLogPath -Repo $next -Folder $RunFolder)
    $next.Status = if ($result.Succeeded) { 'Succeeded' } else { 'Failed' }
    $next.Detail = $result.Detail
    $next.FinishedAt = Get-Date

    $color = if ($result.Succeeded) { 'Green' } else { 'Red' }
    Write-Host "[$($next.Status.ToLower())] $($next.SourceOrg)/$($next.SourceRepo)" -ForegroundColor $color
}

#endregion

#region Main ------------------------------------------------------------------

Assert-Prerequisites

$repos = @(Import-RepoList -Path $RepoListPath)
$runFolder = Join-Path $OutputPath (Get-Date -Format 'yyyyMMdd-HHmmss')
New-Item -ItemType Directory -Path $runFolder -Force | Out-Null

Write-Host 'GEI bulk migration (GitHub Enterprise Cloud)' -ForegroundColor Cyan
Write-Host "  Source         : $(if ($SourceType -eq 'GHES') { $GhesApiUrl } else { 'https://api.github.com' })"
Write-Host '  Destination   : https://api.github.com'
Write-Host "  Storage       : $(if ($SourceType -eq 'GHEC') { 'Not required' } elseif ($UseGitHubStorage) { 'GitHub-owned' } else { 'Azure Blob Storage' })"
Write-Host "  Repositories  : $($repos.Count)"
Write-Host "  Logs          : $runFolder"
Write-Host ''

# Phase 1: queue migrations, keeping at most $ThrottleLimit in flight.
foreach ($repo in $repos) {
    $label = "$($repo.SourceOrg)/$($repo.SourceRepo) -> $($repo.TargetOrg)/$($repo.TargetRepo)"

    if (-not $PSCmdlet.ShouldProcess($label, 'Queue GEI migration')) {
        $repo.Status = 'Skipped'
        $repo.Detail = 'WhatIf'
        continue
    }

    if ($SkipExisting -and (Test-TargetRepoExists -Org $repo.TargetOrg -Repo $repo.TargetRepo)) {
        $repo.Status = 'Skipped'
        $repo.Detail = 'Target repository already exists.'
        Write-Host "[skip]    $label" -ForegroundColor DarkYellow
        continue
    }

    while (@($repos | Where-Object { $_.Status -eq 'Queued' }).Count -ge $ThrottleLimit) {
        Complete-OneMigration -Repos $repos -RunFolder $runFolder
    }

    try {
        $repo.StartedAt = Get-Date
        $repo.MigrationId = Start-RepoMigration -Repo $repo -LogFile (Get-RepoLogPath -Repo $repo -Folder $runFolder)
        $repo.Status = 'Queued'
        Write-Host "[queued]  $label  ($($repo.MigrationId))" -ForegroundColor Gray
    }
    catch {
        $repo.Status = 'Failed'
        $repo.Detail = $_.Exception.Message
        $repo.FinishedAt = Get-Date
        Write-Host "[failed]  $label" -ForegroundColor Red
        Write-Host "          $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Phase 2: drain anything still running.
while (@($repos | Where-Object { $_.Status -eq 'Queued' }).Count -gt 0) {
    Complete-OneMigration -Repos $repos -RunFolder $runFolder
}

$reportPath = Join-Path $runFolder 'migration-report.csv'
$repos | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding utf8

Write-Host ''
Write-Host 'Summary' -ForegroundColor Cyan
$repos | Group-Object Status | ForEach-Object { Write-Host ('  {0,-10} {1}' -f $_.Name, $_.Count) }
Write-Host "  Report     $reportPath"

if (@($repos | Where-Object { $_.Status -eq 'Failed' }).Count -gt 0) { exit 1 }

#endregion
