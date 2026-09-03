# By using the following materials or sample code you agree to be bound by the license terms below
# and the Microsoft Partner Program Agreement the terms of which are incorporated herein by this reference.
# These license terms are an agreement between Microsoft Corporation (or, if applicable based on where you
# are located, one of its affiliates) and you. Any materials (other than sample code) we provide to you
# are for your internal use only. Any sample code is provided for the purpose of illustration only and is
# not intended to be used in a production environment. We grant you a nonexclusive, royalty-free right to
# use and modify the sample code and to reproduce and distribute the object code form of the sample code,
# provided that you agree: (i) to not use Microsoft's name, logo, or trademarks to market your software product
# in which the sample code is embedded; (ii) to include a valid copyright notice on your software product in
# which the sample code is embedded; (iii) to provide on behalf of and for the benefit of your subcontractors
# a disclaimer of warranties, exclusion of liability for indirect and consequential damages and a reasonable
# limitation of liability; and (iv) to indemnify, hold harmless, and defend Microsoft, its affiliates and
# suppliers from and against any third party claims or lawsuits, including attorneys' fees, that arise or result
# from the use or distribution of the sample code.

<#
.SYNOPSIS
    Exports Azure DevOps Server team membership, resolving Distribution List / AD group
    members (including nested groups) so equivalent GitHub teams can be created.

.DESCRIPTION
    Step 1 : Enumerate projects and teams          (Core API)
    Step 2 : Get direct team members               (Core Teams - Members API)
    Step 3 : Resolve group identities              (Identities API / ReadGroupMembers)
    Step 4 : Traverse nested DL membership         (Azure DevOps, AD, or Graph)
    Step 5 : Emit CSV -> ado_team_membership.csv

    NestedExpansion controls how Azure DevOps group identities are expanded:

      AdoApi (default)
        Requires no extra module and no domain connectivity - only the PAT.
        Walks the Azure DevOps Identities API recursively, so nested Azure
        DevOps groups, AD groups, and teams are all expanded with their parent
        group preserved. Reflects membership as Azure DevOps sees it.

      ActiveDirectory
        Uses the RSAT ActiveDirectory module and traverses direct group members.
        Run from a domain-connected Windows machine with permission to read the
        relevant AD groups and users. Azure DevOps-internal groups (project and
        team groups) do not exist in AD and are always expanded through AdoApi.

      Graph
        Uses Microsoft.Graph.Groups and an existing Microsoft Graph sign-in.
        Direct members are traversed recursively so nested Entra ID group paths
        are preserved.

      None
        Uses the Azure DevOps Server internal ReadGroupMembers endpoint and
        returns direct group members only. Nested groups are not expanded.

    If the module required by ActiveDirectory or Graph cannot be loaded, the
    script warns once and falls back to AdoApi instead of failing every group.

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
    Group expansion provider: AdoApi, ActiveDirectory, Graph, or None. The
    default is AdoApi, which needs no additional module.

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
    [ValidateSet('AdoApi','ActiveDirectory','Graph','None')]
    [string] $NestedExpansion = 'AdoApi',
    [string] $OutFile = './ado_team_membership.csv'
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '2026-09-03.1'
$ApiVersion = '6.0'
$Headers = @{
    Authorization = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$Pat"))
    Accept        = 'application/json'
}

Write-Host "Export-AdoTeamDLMembership version $ScriptVersion (NestedExpansion: $NestedExpansion)" -ForegroundColor DarkGray

function Invoke-Ado {
    param([string]$Uri)
    try   { Invoke-RestMethod -Uri $Uri -Headers $Headers -Method Get }
    catch { Write-Warning "GET failed: $Uri  ->  $($_.Exception.Message)"; $null }
}

# ---------------------------------------------------------------------------
# Identity helpers - on-prem returns properties as { "Name": { "$value": ... } }
# ---------------------------------------------------------------------------
function Get-AdoIdentityProperty {
    param($Identity, [string]$Name)
    $value = $Identity.properties.$Name
    if ($null -eq $value)    { return $null }
    if ($value -is [string]) { return $value }
    return [string]$value.'$value'
}

function Get-AdoIdentityName {
    param($Identity)
    foreach ($candidate in @(
        $Identity.providerDisplayName
        $Identity.customDisplayName
        (Get-AdoIdentityProperty $Identity 'Account')
        $Identity.id)) {
        if ($candidate) { return [string]$candidate }
    }
    return '(unknown identity)'
}

function Test-AdoIdentityIsContainer {
    param($Identity)
    if ($null -eq $Identity) { return $false }
    if ($null -ne $Identity.isContainer) { return ($Identity.isContainer -eq $true) }
    $schema = Get-AdoIdentityProperty $Identity 'SchemaClassName'
    if ($schema) { return ($schema -eq 'Group') }
    # Only the ADO-internal SID (S-1-9) implies a group; S-1-5-21 covers both AD users and AD groups.
    return ([string]$Identity.descriptor -match 'GroupScopeType|^Microsoft\.TeamFoundation\.Identity;S-1-9-')
}

# The same person can surface as an ADO GUID, a mail address, a UPN, or DOMAIN\account,
# so each form is indexed - including the bare account name - to match across providers.
function Add-IdentityKey {
    param([System.Collections.Generic.HashSet[string]]$Set, [object[]]$Values)
    foreach ($value in $Values) {
        $text = ([string]$value).Trim()
        if (-not $text) { continue }
        [void]$Set.Add($text)
        if     ($text -match '^[^\\]+\\(.+)$') { [void]$Set.Add($Matches[1]) }
        elseif ($text -match '^([^@]+)@')       { [void]$Set.Add($Matches[1]) }
    }
}

function Test-IdentityKey {
    param([System.Collections.Generic.HashSet[string]]$Set, [object[]]$Values)
    foreach ($value in $Values) {
        $text = ([string]$value).Trim()
        if (-not $text) { continue }
        if ($Set.Contains($text)) { return $true }
        if     ($text -match '^[^\\]+\\(.+)$' -and $Set.Contains($Matches[1])) { return $true }
        elseif ($text -match '^([^@]+)@'       -and $Set.Contains($Matches[1])) { return $true }
    }
    return $false
}

# ---------------------------------------------------------------------------
# Preflight: validate the expansion provider once, not once per group
# ---------------------------------------------------------------------------
$script:ProviderReady = $true
$script:ExpansionFailures = 0
$script:InheritedDirectRows = 0

switch ($NestedExpansion) {
    'ActiveDirectory' {
        try { Import-Module ActiveDirectory -ErrorAction Stop }
        catch {
            $script:ProviderReady = $false
            Write-Warning "ActiveDirectory module could not be loaded: $($_.Exception.Message)"
            Write-Warning "Install RSAT from an elevated session: Add-WindowsCapability -Online -Name 'Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0'"
            Write-Warning 'Falling back to -NestedExpansion AdoApi for this run (no extra module required).'
        }
    }
    'Graph' {
        try { Import-Module Microsoft.Graph.Groups -ErrorAction Stop }
        catch {
            $script:ProviderReady = $false
            Write-Warning "Microsoft.Graph.Groups module could not be loaded: $($_.Exception.Message)"
            Write-Warning 'Install it with: Install-Module Microsoft.Graph.Groups -Scope CurrentUser'
            Write-Warning 'Falling back to -NestedExpansion AdoApi for this run (no extra module required).'
        }
    }
}

function Resolve-ExpansionProvider {
    param($Identity)
    if ($NestedExpansion -eq 'None') { return 'None' }
    # Azure DevOps-internal groups (TFS SID S-1-9-*) have no AD or Entra ID counterpart.
    if ([string]$Identity.descriptor -match 'S-1-9-') { return 'AdoApi' }
    if ($NestedExpansion -in @('ActiveDirectory','Graph') -and -not $script:ProviderReady) { return 'AdoApi' }
    return $NestedExpansion
}

# ---------------------------------------------------------------------------
# Step 4 helper: expand a DL / AD group while preserving membership paths
# ---------------------------------------------------------------------------
$script:GroupCache = @{}

# IIS caps query strings at 2048 bytes by default, so batch by encoded length.
function Get-AdoIdentityBatch {
    param([string[]]$Values, [int]$MaxQueryLength = 1400)

    $batches = [System.Collections.Generic.List[object]]::new()
    $current = [System.Collections.Generic.List[string]]::new()
    $length  = 0

    foreach ($value in $Values) {
        $encoded = [uri]::EscapeDataString([string]$value)
        if ($current.Count -and (($length + $encoded.Length + 3) -gt $MaxQueryLength)) {
            $batches.Add($current.ToArray())
            $current = [System.Collections.Generic.List[string]]::new()
            $length  = 0
        }
        $current.Add($encoded)
        $length += $encoded.Length + 3
    }
    if ($current.Count) { $batches.Add($current.ToArray()) }

    return $batches
}

function Expand-AdoGroupIdentity {
    param([string]$RootId, [string]$RootName)

    if (-not $RootId) { return @() }

    $emitted = [System.Collections.Generic.List[object]]::new()
    $pending = [System.Collections.Generic.Stack[object]]::new()
    $pending.Push([pscustomobject]@{ Id = $RootId; GroupPath = @($RootName); Ancestors = @($RootId) })

    while ($pending.Count -gt 0) {
        $current = $pending.Pop()

        $group = (Invoke-Ado "$CollectionUrl/_apis/identities?identityIds=$($current.Id)&queryMembership=Direct&api-version=$ApiVersion").value |
                 Select-Object -First 1

        # Azure DevOps Server returns both; GUIDs keep the query string far shorter than descriptors.
        $lookupParam = 'identityIds'
        $lookupValues = @($group.memberIds | Where-Object { $_ })
        if (-not $lookupValues.Count) {
            $lookupParam  = 'descriptors'
            $lookupValues = @($group.members | Where-Object { $_ })
        }
        if (-not $lookupValues.Count) { continue }

        foreach ($batch in (Get-AdoIdentityBatch -Values $lookupValues)) {
            $resolved = @((Invoke-Ado "$CollectionUrl/_apis/identities?$lookupParam=$($batch -join ',')&api-version=$ApiVersion").value)

            foreach ($member in $resolved) {
                $memberName = Get-AdoIdentityName $member

                if (Test-AdoIdentityIsContainer $member) {
                    $childId = [string]$member.id
                    if (-not $childId -or ($childId -in @($current.Ancestors))) {
                        Write-Warning "Skipping cyclic or unresolvable group membership: $((@($current.GroupPath) + $memberName) -join ' > ')"
                        continue
                    }
                    $pending.Push([pscustomobject]@{
                        Id        = $childId
                        GroupPath = @($current.GroupPath) + $memberName
                        Ancestors = @($current.Ancestors) + $childId
                    })
                    continue
                }

                $mail    = Get-AdoIdentityProperty $member 'Mail'
                $account = Get-AdoIdentityProperty $member 'Account'
                $emitted.Add([pscustomobject]@{
                    DisplayName     = $memberName
                    Mail            = $mail
                    UserId          = if ($mail) { $mail } elseif ($account) { $account } else { $memberName }
                    IdentityId      = [string]$member.id
                    Account         = $account
                    NestedGroup     = @($current.GroupPath)[-1]
                    SourceGroupPath = @($current.GroupPath) -join ' > '
                    NestingDepth    = (@($current.GroupPath).Count - 1)
                })
            }
        }
    }

    return $emitted.ToArray()
}

function Expand-GroupMembers {
    param([string]$GroupName, [string]$GroupIdentityId, [string]$Provider = $NestedExpansion)

    $key = "$Provider|$GroupName|$GroupIdentityId"
    if ($script:GroupCache.ContainsKey($key)) { return $script:GroupCache[$key] }

    $people = @()

    switch ($Provider) {

        'AdoApi' {
            $people = @(Expand-AdoGroupIdentity -RootId $GroupIdentityId -RootName $GroupName)
        }

        'ActiveDirectory' {
            # Strips the "DOMAIN\" / "(Cognizant)" decoration ADO adds to the display name
            $sam = ($GroupName -replace '^.*\\', '' -replace '\s*\(.*\)\s*$', '').Trim()
            try {
                $people = @(& {
                    $pending = [System.Collections.Generic.Stack[object]]::new()
                    $pending.Push([pscustomobject]@{
                        GroupId   = $sam
                        GroupPath = @($GroupName)
                        Ancestors = @($sam.ToLowerInvariant())
                    })

                    while ($pending.Count -gt 0) {
                        $current = $pending.Pop()

                        foreach ($member in @(Get-ADGroupMember -Identity $current.GroupId)) {
                            if ($member.objectClass -ieq 'group') {
                                $childId = if ($member.DistinguishedName) {
                                    [string]$member.DistinguishedName
                                } else {
                                    [string]$member.SamAccountName
                                }
                                $childKey = if ($member.SamAccountName) {
                                    ([string]$member.SamAccountName).ToLowerInvariant()
                                } else {
                                    $childId.ToLowerInvariant()
                                }
                                $childName = if ($member.Name) {
                                    [string]$member.Name
                                } else {
                                    [string]$member.SamAccountName
                                }

                                if ($childKey -in @($current.Ancestors)) {
                                    $cyclePath = (@($current.GroupPath) + $childName) -join ' > '
                                    Write-Warning "Skipping cyclic AD group membership: $cyclePath"
                                    continue
                                }

                                $pending.Push([pscustomobject]@{
                                    GroupId   = $childId
                                    GroupPath = @($current.GroupPath) + $childName
                                    Ancestors = @($current.Ancestors) + $childKey
                                })
                                continue
                            }

                            if ($member.objectClass -ine 'user') { continue }

                            $userId = if ($member.DistinguishedName) {
                                [string]$member.DistinguishedName
                            } else {
                                [string]$member.SamAccountName
                            }
                            $u = Get-ADUser -Identity $userId -Properties mail, DisplayName, UserPrincipalName
                            [pscustomobject]@{
                                DisplayName     = $u.DisplayName
                                Mail            = $u.mail
                                UserId          = if ($u.UserPrincipalName) { $u.UserPrincipalName } else { $u.SamAccountName }
                                IdentityId      = $null
                                Account         = [string]$u.SamAccountName
                                NestedGroup     = @($current.GroupPath)[-1]
                                SourceGroupPath = @($current.GroupPath) -join ' > '
                                NestingDepth    = (@($current.GroupPath).Count - 1)
                            }
                        }
                    }
                })
            } catch {
                $script:ExpansionFailures++
                Write-Warning "AD expansion failed for '$sam': $($_.Exception.Message)"
            }
        }

        'Graph' {
            try {
                $escapedGroupName = $GroupName.Replace("'", "''")
                $g = Get-MgGroup -Filter "displayName eq '$escapedGroupName'" -Top 1
                if ($g) {
                    $people = @(& {
                        $pending = [System.Collections.Generic.Stack[object]]::new()
                        $pending.Push([pscustomobject]@{
                            GroupId   = [string]$g.Id
                            GroupPath = @($GroupName)
                            Ancestors = @([string]$g.Id)
                        })

                        while ($pending.Count -gt 0) {
                            $current = $pending.Pop()

                            foreach ($member in @(Get-MgGroupMember -GroupId $current.GroupId -All)) {
                                $additional = $member.AdditionalProperties
                                $odataType = if ($additional) { [string]$additional['@odata.type'] } else { '' }

                                if ($odataType -eq '#microsoft.graph.group') {
                                    $childId = [string]$member.Id
                                    $childName = if ($additional.displayName) {
                                        [string]$additional.displayName
                                    } else {
                                        $childId
                                    }

                                    if ($childId -in @($current.Ancestors)) {
                                        $cyclePath = (@($current.GroupPath) + $childName) -join ' > '
                                        Write-Warning "Skipping cyclic Entra ID group membership: $cyclePath"
                                        continue
                                    }

                                    $pending.Push([pscustomobject]@{
                                        GroupId   = $childId
                                        GroupPath = @($current.GroupPath) + $childName
                                        Ancestors = @($current.Ancestors) + $childId
                                    })
                                    continue
                                }

                                if ($odataType -ne '#microsoft.graph.user') { continue }

                                $principalName = [string]$additional.userPrincipalName
                                [pscustomobject]@{
                                    DisplayName     = [string]$additional.displayName
                                    Mail            = [string]$additional.mail
                                    UserId          = if ($principalName) { $principalName } else { [string]$member.Id }
                                    IdentityId      = [string]$member.Id
                                    Account         = $principalName
                                    NestedGroup     = @($current.GroupPath)[-1]
                                    SourceGroupPath = @($current.GroupPath) -join ' > '
                                    NestingDepth    = (@($current.GroupPath).Count - 1)
                                }
                            }
                        }
                    })
                }
            } catch {
                $script:ExpansionFailures++
                Write-Warning "Graph expansion failed for '$GroupName': $($_.Exception.Message)"
            }
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
                                      DisplayName     = $_.DisplayName
                                      Mail            = $_.MailAddress
                                      UserId          = $_.AccountName
                                      IdentityId      = [string]$_.TeamFoundationId
                                      Account         = [string]$_.AccountName
                                      NestedGroup     = $GroupName
                                      SourceGroupPath = $GroupName
                                      NestingDepth    = 0
                                  }
                              }
                }
            }
        }
    }

    $people = @($people | Sort-Object UserId, SourceGroupPath -Unique)
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

        $slug = ($t.name -replace '[^a-zA-Z0-9]+','-').ToLower().Trim('-')

        # A team is itself a group identity, so its Direct membership is the authoritative
        # list of what was added to the team - anything else arrived through a group.
        $teamDirectIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $teamIdentity = (Invoke-Ado "$CollectionUrl/_apis/identities?identityIds=$($t.id)&queryMembership=Direct&api-version=$ApiVersion").value |
                        Select-Object -First 1
        foreach ($memberId in @($teamIdentity.memberIds | Where-Object { $_ })) { [void]$teamDirectIds.Add([string]$memberId) }

        # This endpoint may return a DL as a group identity AND the users inside it.
        $members = (Invoke-Ado "$CollectionUrl/_apis/projects/$($p.id)/teams/$($t.id)/members?api-version=$ApiVersion&`$top=1000").value

        $groupEntries = [System.Collections.Generic.List[object]]::new()
        $userEntries  = [System.Collections.Generic.List[object]]::new()

        foreach ($m in $members) {
            $id = $m.identity

            # Resolve the identity to determine whether it's a container (group/DL) or a user
            $ident = (Invoke-Ado "$CollectionUrl/_apis/identities?identityIds=$($id.id)&queryMembership=Expanded&api-version=$ApiVersion").value | Select-Object -First 1
            $isGroup = Test-AdoIdentityIsContainer $ident

            if ($isGroup) { $groupEntries.Add([pscustomobject]@{ Member = $m; Identity = $ident }) }
            else          { $userEntries.Add([pscustomobject]@{ Member = $m; Identity = $ident }) }
        }

        # Groups are expanded first so a user inherited from a DL is not also reported as direct.
        $inheritedKeys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

        Write-Verbose "Team '$($t.name)': $($teamDirectIds.Count) direct identity/identities, $($groupEntries.Count) group(s), $($userEntries.Count) user entry/entries."

        foreach ($entry in $groupEntries) {
            $m    = $entry.Member
            $id   = $m.identity
            $name = $id.displayName

            $provider = Resolve-ExpansionProvider -Identity $entry.Identity
            foreach ($u in (Expand-GroupMembers -GroupName $name -GroupIdentityId $id.id -Provider $provider)) {
                $membershipType = if ($u.NestingDepth -gt 0) { 'Nested group member' } else { 'Direct group member' }
                $rows.Add([pscustomobject]@{
                    Project           = $p.name
                    AdoTeam           = $t.name
                    SourceType        = 'DL/Group (expanded)'
                    SourceGroup       = $name
                    NestedGroup       = $u.NestedGroup
                    MembershipType    = $membershipType
                    NestingDepth      = $u.NestingDepth
                    MemberDisplayName = $u.DisplayName
                    MemberEmail       = $u.Mail
                    MemberUserId      = $u.UserId
                    IsTeamAdmin       = $m.isTeamAdmin
                    SuggestedGitHubTeam = $slug
                })
                Add-IdentityKey -Set $inheritedKeys -Values @($u.IdentityId, $u.Mail, $u.UserId, $u.Account, $u.DisplayName)
            }
        }

        foreach ($entry in $userEntries) {
            $m    = $entry.Member
            $id   = $m.identity
            $name = $id.displayName

            $candidateKeys = @(
                $id.id
                $id.uniqueName
                (Get-AdoIdentityProperty $entry.Identity 'Mail')
                (Get-AdoIdentityProperty $entry.Identity 'Account')
                $name
            )

            # Prefer the team's own Direct membership; fall back to key matching only when unavailable.
            $isDirect = if ($teamDirectIds.Count) {
                $teamDirectIds.Contains([string]$id.id)
            } else {
                -not (Test-IdentityKey -Set $inheritedKeys -Values $candidateKeys)
            }

            if (-not $isDirect) {
                $script:InheritedDirectRows++
                Write-Verbose "  Inherited through a group, not listed as direct: $name <$($id.uniqueName)>"
                continue
            }

            $rows.Add([pscustomobject]@{
                Project           = $p.name
                AdoTeam           = $t.name
                SourceType        = 'Direct user'
                SourceGroup       = ''
                NestedGroup       = ''
                MembershipType    = 'Direct ADO team member'
                NestingDepth      = 0
                MemberDisplayName = $name
                MemberEmail       = $id.uniqueName
                MemberUserId      = $id.uniqueName
                IsTeamAdmin       = $m.isTeamAdmin
                SuggestedGitHubTeam = $slug
            })
        }
    }
}

$exportRows = @($rows | Sort-Object Project, AdoTeam, SourceGroup, NestedGroup, MemberUserId -Unique)
$exportRows | Export-Csv -Path $OutFile -NoTypeInformation -Encoding UTF8

if ($script:InheritedDirectRows -gt 0) {
    Write-Host "Suppressed $($script:InheritedDirectRows) duplicate row(s) for users who reached a team through a group." -ForegroundColor DarkGray
}

if ($script:ExpansionFailures -gt 0) {
    Write-Warning "$($script:ExpansionFailures) group(s) could not be expanded - the CSV is INCOMPLETE. Resolve the warnings above before using it for migration."
    Write-Host "`nExported $($exportRows.Count) rows -> $OutFile" -ForegroundColor Yellow
}
else {
    Write-Host "`nExported $($exportRows.Count) rows -> $OutFile" -ForegroundColor Green
}
Write-Host "Next: feed MemberEmail + SuggestedGitHubTeam into the GitHub REST API:" -ForegroundColor Yellow
Write-Host "  POST /orgs/{org}/teams" -ForegroundColor DarkGray
Write-Host "  PUT  /orgs/{org}/teams/{team_slug}/memberships/{username}" -ForegroundColor DarkGray
