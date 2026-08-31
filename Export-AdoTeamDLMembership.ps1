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
    Exports Azure DevOps Server team membership, resolving Distribution List / AD group
    members (including nested groups) so equivalent GitHub teams can be created.

.DESCRIPTION
    Step 1 : Enumerate projects and teams          (Core API)
    Step 2 : Get direct team members               (Core Teams - Members API)
    Step 3 : Resolve group identities              (Identities API / ReadGroupMembers)
    Step 4 : Flatten nested DL membership          (AD -Recursive  OR  Graph /transitiveMembers)
    Step 5 : Emit CSV -> ado_team_membership.csv

    NestedExpansion controls how Azure DevOps group identities are expanded:

      ActiveDirectory (default)
        Uses the RSAT ActiveDirectory module and Get-ADGroupMember -Recursive.
        Run from a domain-connected Windows machine with permission to read the
        relevant AD groups and users.

      Graph
        Uses Microsoft.Graph.Groups and an existing Microsoft Graph sign-in.
        Get-MgGroupTransitiveMember recursively returns nested Entra ID members.

      None
        Requires no AD or Graph module. Uses the Azure DevOps Server internal
        ReadGroupMembers endpoint and returns direct group members only. Nested
        groups are not expanded. This is the recommended mode for a first test.

    The script only reads Azure DevOps, AD, or Graph data and writes a CSV. It
    does not create GitHub teams or add GitHub users.

.PARAMETER CollectionUrl
    Full Azure DevOps Server collection URL, without a trailing slash. Example:
    https://devops.example.com/DefaultCollection

.PARAMETER Pat
    Azure DevOps Server personal access token with Identity (Read) and Project
    and Team (Read) access. Prompt for this value at runtime instead of placing
    it directly in the command or saving it in the script.

.PARAMETER Projects
    Optional exact project-name filter. Omit it to process every visible project.
    Supply multiple projects as a PowerShell array, for example:
    -Projects 'Payments','Shared Services'

.PARAMETER NestedExpansion
    Group expansion provider: ActiveDirectory, Graph, or None. The default is
    ActiveDirectory.

.PARAMETER OutFile
    Destination CSV path. A relative path is resolved from the current working
    directory, not from the directory containing this script. Use an absolute
    path when the output location matters.

.EXAMPLE
    PS> Get-Help .\Export-AdoTeamDLMembership.ps1 -Full

    Displays this help, including parameter descriptions and examples.

.EXAMPLE
    PS> $securePat = Read-Host 'Azure DevOps PAT' -AsSecureString
    PS> $pat = [System.Net.NetworkCredential]::new('', $securePat).Password
    PS> $outFile = Join-Path $PWD 'ado_team_membership.csv'
    PS> try {
    >>     .\Export-AdoTeamDLMembership.ps1 `
    >>         -CollectionUrl 'https://devops.example.com/DefaultCollection' `
    >>         -Pat $pat `
    >>         -Projects 'Known Small Project' `
    >>         -NestedExpansion None `
    >>         -OutFile $outFile
    >> }
    >> finally {
    >>     Remove-Variable pat, securePat -ErrorAction SilentlyContinue
    >> }

    Performs the recommended first run against one known project. This mode
    exports direct users and direct members of Azure DevOps groups, but it does
    not recurse into nested groups.

.EXAMPLE
    PS> Import-Module ActiveDirectory
    PS> Get-ADGroup 'Known AD Group' | Out-Null
    PS> $pat = Read-Host 'Azure DevOps PAT' -MaskInput
    PS> .\Export-AdoTeamDLMembership.ps1 `
    >>     -CollectionUrl 'https://devops.example.com/DefaultCollection' `
    >>     -Pat $pat `
    >>     -Projects 'Payments','Shared Services' `
    >>     -NestedExpansion ActiveDirectory `
    >>     -OutFile 'C:\Exports\ado_team_membership.csv'
    PS> Remove-Variable pat

    Recursively expands on-premises AD groups. On Windows 10 or 11, install the
    RSAT Active Directory capability first from an elevated PowerShell session:

    Add-WindowsCapability -Online `
        -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'

.EXAMPLE
    PS> Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
    PS> Install-Module Microsoft.Graph.Groups -Scope CurrentUser
    PS> Connect-MgGraph -Scopes 'Group.Read.All','User.Read.All'
    PS> $pat = Read-Host 'Azure DevOps PAT' -MaskInput
    PS> .\Export-AdoTeamDLMembership.ps1 `
    >>     -CollectionUrl 'https://devops.example.com/DefaultCollection' `
    >>     -Pat $pat `
    >>     -Projects 'Cloud Platform' `
    >>     -NestedExpansion Graph `
    >>     -OutFile 'C:\Exports\ado_team_membership.csv'
    PS> Remove-Variable pat
    PS> Disconnect-MgGraph

    Recursively expands Entra ID groups through Microsoft Graph. The requested
    Graph permissions may require tenant administrator consent.

.EXAMPLE
    PS> $rows = @(Import-Csv -LiteralPath 'C:\Exports\ado_team_membership.csv')
    PS> if ($rows.Count -eq 0) { throw 'No membership rows were exported.' }
    PS> $rows | Group-Object Project, AdoTeam, SourceType |
    >>     Select-Object Count, Name | Format-Table -AutoSize
    PS> $rows | Select-Object -First 10 | Format-Table -AutoSize

    Verifies that the CSV is readable and nonempty, then summarizes and samples
    the exported membership. Compare several rows with the Azure DevOps team UI
    and the authoritative AD or Entra ID group membership.

.NOTES
    Azure DevOps Server (on-prem). PAT needs: Identity (Read), Project & Team (Read).
    Nested expansion requires either RSAT ActiveDirectory module or Microsoft.Graph module.

    Before the first run:
      1. Review the script and confirm CollectionUrl points to the intended server.
      2. If Windows marks the reviewed download as blocked, run:
           Unblock-File -LiteralPath .\Export-AdoTeamDLMembership.ps1
         Do not bypass an organization-enforced execution policy; contact the
         administrator if Unblock-File is not permitted.
      3. Verify API access with one known project and NestedExpansion None.
      4. Treat every "GET failed" or "expansion failed" warning as an incomplete
         export, even if the script reaches its final "Exported" message.

    The API queries request at most 1000 projects, teams, or members and do not
    currently follow continuation tokens. Validate completeness before using the
    CSV for migration when any collection may exceed that limit.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $CollectionUrl,          # e.g. https://devops.cognizant.com/DefaultCollection
    [Parameter(Mandatory)] [string] $Pat,
    [string[]] $Projects,                                     # optional filter; omit for all
    [ValidateSet('ActiveDirectory','Graph','None')]
    [string] $NestedExpansion = 'ActiveDirectory',
    [string] $OutFile = './ado_team_membership.csv'
)

$ErrorActionPreference = 'Stop'
$ApiVersion = '6.0'
$Headers = @{
    Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat"))
    Accept        = 'application/json'
}

function Invoke-Ado {
    param([string]$Uri)
    try   { Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get }
    catch { Write-Warning "GET failed: $Uri  ->  $($_.Exception.Message)"; $null }
}

# ---------------------------------------------------------------------------
# Step 4 helper: flatten a DL / AD group into individual users
# ---------------------------------------------------------------------------
$script:GroupCache = @{}

function Expand-GroupMembers {
    param([string]$GroupName, [string]$GroupIdentityId)

    $key = "$GroupName|$GroupIdentityId"
    if ($script:GroupCache.ContainsKey($key)) { return $script:GroupCache[$key] }

    $people = @()

    switch ($NestedExpansion) {

        'ActiveDirectory' {
            # Strips the "DOMAIN\" / "(Cognizant)" decoration ADO adds to the display name
            $sam = ($GroupName -replace '^.*\\', '' -replace '\s*\(.*\)\s*$', '').Trim()
            try {
                Import-Module ActiveDirectory -ErrorAction Stop
                # -Recursive flattens ALL nested groups in a single call
                $people = Get-ADGroupMember -Identity $sam -Recursive |
                          Where-Object { $_.objectClass -eq 'user' } |
                          ForEach-Object {
                              $u = Get-ADUser $_.SamAccountName -Properties mail, DisplayName
                              [pscustomobject]@{
                                  DisplayName = $u.DisplayName
                                  Mail        = $u.mail
                                  UserId      = $u.UserPrincipalName
                              }
                          }
            } catch { Write-Warning "AD expansion failed for '$sam': $($_.Exception.Message)" }
        }

        'Graph' {
            # Entra ID: /groups/{id}/transitiveMembers returns a flat list of all nested members
            try {
                Import-Module Microsoft.Graph.Groups -ErrorAction Stop
                $g = Get-MgGroup -Filter "displayName eq '$GroupName'" -Top 1
                if ($g) {
                    $people = Get-MgGroupTransitiveMember -GroupId $g.Id -All |
                              Where-Object { $_.AdditionalProperties['@odata.type'] -eq '#microsoft.graph.user' } |
                              ForEach-Object {
                                  [pscustomobject]@{
                                      DisplayName = $_.AdditionalProperties.displayName
                                      Mail        = $_.AdditionalProperties.mail
                                      UserId      = $_.AdditionalProperties.userPrincipalName
                                  }
                              }
                }
            } catch { Write-Warning "Graph expansion failed for '$GroupName': $($_.Exception.Message)" }
        }

        'None' {
            # Fall back to the ADO Server internal endpoint the web UI itself calls.
            # NOTE: direct members only - it does NOT recurse into nested groups.
            if ($GroupIdentityId) {
                $uri = "$CollectionUrl/_api/_identity/ReadGroupMembers?scope=$GroupIdentityId&readMembers=true&api-version=$ApiVersion"
                $r = Invoke-Ado $uri
                if ($r -and $r.identities) {
                    $people = $r.identities |
                              Where-Object { -not $_.IsContainer } |
                              ForEach-Object {
                                  [pscustomobject]@{
                                      DisplayName = $_.DisplayName
                                      Mail        = $_.MailAddress
                                      UserId      = $_.AccountName
                                  }
                              }
                }
            }
        }
    }

    $people = @($people | Sort-Object UserId -Unique)
    $script:GroupCache[$key] = $people
    return $people
}

# ---------------------------------------------------------------------------
# Step 1-3: walk projects -> teams -> members
# ---------------------------------------------------------------------------
$rows = [System.Collections.Generic.List[object]]::new()

$projList = (Invoke-Ado "$CollectionUrl/_apis/projects?api-version=$ApiVersion&`$top=1000").value
if ($Projects) { $projList = $projList | Where-Object { $_.name -in $Projects } }

foreach ($p in $projList) {
    Write-Host "Project: $($p.name)" -ForegroundColor Cyan

    $teams = (Invoke-Ado "$CollectionUrl/_apis/projects/$($p.id)/teams?api-version=$ApiVersion&`$top=1000").value
    foreach ($t in $teams) {
        Write-Host "  Team: $($t.name)"

        # Direct members only - a DL comes back as ONE group identity, not its people
        $members = (Invoke-Ado "$CollectionUrl/_apis/projects/$($p.id)/teams/$($t.id)/members?api-version=$ApiVersion&`$top=1000").value
        foreach ($m in $members) {

            $id   = $m.identity
            $name = $id.displayName

            # Resolve the identity to determine whether it's a container (group/DL) or a user
            $ident = (Invoke-Ado "$CollectionUrl/_apis/identities?identityIds=$($id.id)&queryMembership=Expanded&api-version=$ApiVersion").value | Select-Object -First 1
            $isGroup = $false
            if ($ident) {
                $isGroup = ($ident.properties.SchemaClassName.'$value' -eq 'Group') -or
                           ($ident.descriptor -match 'GroupScopeType|^Microsoft\.TeamFoundation\.Identity;S-1-5-') -or
                           ($ident.isContainer -eq $true)
            }

            if ($isGroup) {
                foreach ($u in (Expand-GroupMembers -GroupName $name -GroupIdentityId $id.id)) {
                    $rows.Add([pscustomobject]@{
                        Project          = $p.name
                        AdoTeam          = $t.name
                        SourceType       = 'DL/Group (expanded)'
                        SourceGroup      = $name
                        MemberDisplayName= $u.DisplayName
                        MemberEmail      = $u.Mail
                        MemberUserId     = $u.UserId
                        IsTeamAdmin      = $m.isTeamAdmin
                        SuggestedGitHubTeam = ($t.name -replace '[^a-zA-Z0-9]+','-').ToLower().Trim('-')
                    })
                }
            }
            else {
                $rows.Add([pscustomobject]@{
                    Project          = $p.name
                    AdoTeam          = $t.name
                    SourceType       = 'Direct user'
                    SourceGroup      = ''
                    MemberDisplayName= $name
                    MemberEmail      = $id.uniqueName
                    MemberUserId     = $id.uniqueName
                    IsTeamAdmin      = $m.isTeamAdmin
                    SuggestedGitHubTeam = ($t.name -replace '[^a-zA-Z0-9]+','-').ToLower().Trim('-')
                })
            }
        }
    }
}

$rows | Sort-Object Project, AdoTeam, MemberUserId -Unique | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8
Write-Host "`nExported $($rows.Count) rows -> $OutFile" -ForegroundColor Green
Write-Host "Next: feed MemberEmail + SuggestedGitHubTeam into the GitHub REST API:" -ForegroundColor Yellow
Write-Host "  POST /orgs/{org}/teams" -ForegroundColor DarkGray
Write-Host "  PUT  /orgs/{org}/teams/{team_slug}/memberships/{username}" -ForegroundColor DarkGray
