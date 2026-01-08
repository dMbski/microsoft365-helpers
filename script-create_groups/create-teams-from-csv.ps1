#functions: create m365 groups, add owner, add members from souce group, create team with parameters.
#data are from csv file with headers: Owner,MailNickName,SectionNumber,Description,CourseName,DisplayName,SourceM365Group
#										m		m			n				m			n			m			m
#		m - mandatory fields
#		example						mathteacher@sp.edu,class7a-math,7a,Math in 7a,Math,Class 7a Math,class7a-all@sp.edu

# Paths
$csvDataPath  = "teams.csv"
$csvErrorPath = "teams-errors.csv"

function Write-ErrorCsv {
    param(
        [string]$Owner,
        [string]$MailNickName,
        [string]$DisplayName,
        [string]$SourceM365Group,
        [string]$ErrorMessage
    )

    # Append a single error line to the error CSV file
    $line = '"' + ($Owner          -replace '"','""') + '",' +
            '"' + ($MailNickName   -replace '"','""') + '",' +
            '"' + ($DisplayName    -replace '"','""') + '",' +
            '"' + ($SourceM365Group-replace '"','""') + '",' +
            '"' + ($ErrorMessage   -replace '"','""') + '"'

    Add-Content -Path $csvErrorPath -Value $line
}

# Check if input CSV file exists
if (-not (Test-Path -Path $csvDataPath)) {
    Write-Host "Input CSV file does not exist: $csvDataPath"
    exit 1
}
Write-Host "Input CSV file found: $csvDataPath"

# Load data from CSV
$rows = Import-Csv -Path $csvDataPath

# Always recreate error CSV with header
"Owner,MailNickName,DisplayName,SourceM365Group,Error" |
    Out-File -FilePath $csvErrorPath -Encoding UTF8 -Force

# Connect to Microsoft Graph only after CSV validation
Import-Module Microsoft.Graph

Write-Host "Connecting to Microsoft Graph..."
Connect-MgGraph -Scopes `
    "Team.Create",`
    "TeamSettings.ReadWrite.All",`
    "Group.ReadWrite.All",`
    "Directory.Read.All",`
    "User.Read.All"

$ctx = Get-MgContext
if (-not $ctx.Account) {
    Write-Error "No Graph context – Connect-MgGraph failed."
    exit 1
}
Write-Host "Microsoft Graph connection OK, starting CSV processing..."

foreach ($row in $rows) {

    # Read row fields from CSV
    $ownerUpn        = $row.Owner
    $mailNickname    = $row.MailNickName
    $sectionNumber   = $row.SectionNumber
    $description     = $row.Description
    $courseName      = $row.CourseName
    $displayName     = $row.DisplayName
    $sourceMail      = $row.SourceM365Group   # full mail address of source group

    Write-Host ""
    Write-Host "Processing row for group '$displayName' (source: '$sourceMail')"

    # ----- VALIDATION: required fields -----
    if ([string]::IsNullOrWhiteSpace($displayName) -or
        [string]::IsNullOrWhiteSpace($mailNickname)) {

        Write-Host "Missing DisplayName or MailNickName -> skipping row"
        Write-ErrorCsv -Owner $ownerUpn `
                       -MailNickName $mailNickname `
                       -DisplayName $displayName `
                       -SourceM365Group $sourceMail `
                       -ErrorMessage "Missing DisplayName or MailNickName"
        continue
    }

    # ----- VALIDATION: owner must exist (if provided) -----
    $ownerId = $null
    if (-not [string]::IsNullOrWhiteSpace($ownerUpn)) {
        try {
            # Get user object by UPN or Id
            $owner = Get-MgUser -UserId $ownerUpn -ErrorAction Stop
            $ownerId = $owner.Id
            Write-Host "Owner found: $ownerUpn"
        }
        catch {
            Write-Host "Owner NOT found: $ownerUpn"
            Write-ErrorCsv -Owner $ownerUpn `
                           -MailNickName $mailNickname `
                           -DisplayName $displayName `
                           -SourceM365Group $sourceMail `
                           -ErrorMessage "Owner not found: $ownerUpn"
            continue
        }
    } else {
        Write-Host "No owner specified in CSV"
    }

    # ----- VALIDATION: source M365 group by mail (if provided) -----
    $sourceGroup      = $null
    $sourceGroupId    = $null
    $sourceMembers    = @()   # list of members from source group

    if (-not [string]::IsNullOrWhiteSpace($sourceMail)) {
        try {
            $filter = "mail eq '$sourceMail'"
            Write-Host "Looking for source group $sourceMail"

            $sourceGroups = Get-MgGroup -Filter $filter -ErrorAction Stop

            # Handle both single object and collection
            if ($sourceGroups -is [System.Array]) {
                $groupCount = $sourceGroups.Count
            } elseif ($null -ne $sourceGroups) {
                $groupCount  = 1
                $sourceGroups = @($sourceGroups)
            } else {
                $groupCount = 0
            }

            Write-Host "Found groups count: $groupCount"

            if ($groupCount -ge 1) {
                $sourceGroup   = $sourceGroups[0]
                $sourceGroupId = $sourceGroup.Id
                Write-Host "Source group found by mail: $sourceMail (Id: $sourceGroupId)"

                # Load members of source group here (for debugging and reuse)
                if ($sourceGroupId) {
                    try {
                        $sourceMembers = Get-MgGroupMember -GroupId $sourceGroupId -All
                        $memberCount   = if ($sourceMembers) { $sourceMembers.Count } else { 0 }
                        Write-Host "Source group member count: $memberCount"
                    }
                    catch {
                        Write-Host "Failed to read members from source group: $($_.Exception.Message)"
                        Write-ErrorCsv -Owner $ownerUpn `
                                       -MailNickName $mailNickname `
                                       -DisplayName $displayName `
                                       -SourceM365Group $sourceMail `
                                       -ErrorMessage "Failed to read members from source group: $($_.Exception.Message)"
                        # still can create group without copying members
                    }
                } else {
                    Write-Host "SourceGroupId is empty after search – skipping member read"
                    Write-ErrorCsv -Owner $ownerUpn `
                                   -MailNickName $mailNickname `
                                   -DisplayName $displayName `
                                   -SourceM365Group $sourceMail `
                                   -ErrorMessage "SourceGroupId empty after search by mail: $sourceMail"
                }
            }
            else {
                Write-Host "Source group NOT found by mail: $sourceMail"
                Write-ErrorCsv -Owner $ownerUpn `
                               -MailNickName $mailNickname `
                               -DisplayName $displayName `
                               -SourceM365Group $sourceMail `
                               -ErrorMessage "SourceM365Group not found by mail: $sourceMail"
                continue
            }
        }
        catch {
            Write-Host "Error searching source group by mail: $($_.Exception.Message)"
            Write-ErrorCsv -Owner $ownerUpn `
                           -MailNickName $mailNickname `
                           -DisplayName $displayName `
                           -SourceM365Group $sourceMail `
                           -ErrorMessage "Error searching SourceM365Group by mail: $($_.Exception.Message)"
            continue
        }
    } else {
        Write-Host "No source group mail specified in CSV"
    }

    # ----- STEP 1: Ensure target M365 group exists (Unified, Private) -----
    $group = $null

    # First: check if group already exists in tenant by mailNickname
    try {
        $existingFilter = "mailNickname eq '$mailNickname'"
        Write-Host "Checking if target group already exists $mailNickname"

        $existingGroups = Get-MgGroup -Filter $existingFilter -ErrorAction Stop

        if ($existingGroups -is [System.Array]) {
            $existingCount = $existingGroups.Count
        } elseif ($null -ne $existingGroups) {
            $existingCount  = 1
            $existingGroups = @($existingGroups)
        } else {
            $existingCount = 0
        }

        Write-Host "Existing target groups count: $existingCount"

        if ($existingCount -gt 0) {
            $group = $existingGroups[0]
            Write-Host "Target group already exists: $($group.DisplayName) ($($group.Id))"
        } else {
            Write-Host "Target group does not exist, will create new one..."
        }
    }
    catch {
        Write-Host "Error while checking existing group: $($_.Exception.Message)"
        # in case of error, we still try to create a new group
    }

    # If group was not found, create it now
    if (-not $group) {
        try {
            $groupParams = @{
                DisplayName     = $displayName
                MailNickname    = $mailNickname
                Description     = $description
                GroupTypes      = @("Unified")
                MailEnabled     = $true
                SecurityEnabled = $false
                Visibility      = "Private"
            }

            $group = New-MgGroup -BodyParameter $groupParams
            Write-Host "Created group: $($group.DisplayName) ($($group.Id))"
        }
        catch {
            Write-Host "Group creation failed: $($_.Exception.Message)"
            Write-ErrorCsv -Owner $ownerUpn `
                           -MailNickName $mailNickname `
                           -DisplayName $displayName `
                           -SourceM365Group $sourceMail `
                           -ErrorMessage "Group creation failed: $($_.Exception.Message)"
            continue
        }
    }

    if (-not $group -or -not $group.Id) {
        Write-Host "Group reference is empty -> skipping row"
        Write-ErrorCsv -Owner $ownerUpn `
                       -MailNickName $mailNickname `
                       -DisplayName $displayName `
                       -SourceM365Group $sourceMail `
                       -ErrorMessage "Group reference empty after existence check / creation."
        continue
    }

    # ----- STEP 2: Assign owner to the new/existing group (if resolved) -----
    if ($ownerId) {
        try {
            New-MgGroupOwnerByRef -GroupId $group.Id -BodyParameter @{
                "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$ownerId"
            }
            Write-Host "Added owner $ownerUpn to group $($group.DisplayName)"
        }
        catch {
            Write-Host "Failed to add owner: $($_.Exception.Message)"
            Write-ErrorCsv -Owner $ownerUpn `
                           -MailNickName $mailNickname `
                           -DisplayName $displayName `
                           -SourceM365Group $sourceMail `
                           -ErrorMessage "Failed to add owner: $($_.Exception.Message)"
            # owner failure does NOT stop member copy or team creation
        }
    } else {
        Write-Host "Owner not set for target group"
    }

    # ----- STEP 3: Copy members from source group (if we have them) -----
    if ($sourceGroupId -and $sourceMembers.Count -gt 0) {
        Write-Host "Copying $($sourceMembers.Count) members from '$sourceMail' to '$displayName'..."

        $idsPreview = ($sourceMembers | Select-Object -First 5 -ExpandProperty Id) -join ", "
        Write-Host "First member Ids: $idsPreview"

        foreach ($member in $sourceMembers) {
            try {
                Write-Host "  Adding member $($member.Id) to group $($group.DisplayName)..."
                New-MgGroupMemberByRef -GroupId $group.Id -BodyParameter @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($member.Id)"
                }
            }
            catch {
                $addErr = $_.Exception.Message
                if ($addErr -like "*One or more added object references already exist*") {
                    Write-Host "  Member $($member.Id) is already in target group – skipping."
                }
                else {
                    Write-Host "  Failed to add member $($member.Id): $addErr"
                    Write-ErrorCsv -Owner $ownerUpn `
                                   -MailNickName $mailNickname `
                                   -DisplayName $displayName `
                                   -SourceM365Group $sourceMail `
                                   -ErrorMessage "Failed to add member $($member.Id) from source group: $addErr"
                }
            }
        }

        Write-Host "Finished copying members to '$displayName'"
    }
    else {
        Write-Host "No members to copy from source group (Id: $sourceGroupId, Count: $($sourceMembers.Count))"
    }

    Start-Sleep -Seconds 5

    # ----- STEP 4: Enable Teams for the group with specific settings -----
    try {
        $teamParams = @{
            "template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('standard')"
            "group@odata.bind"    = "https://graph.microsoft.com/v1.0/groups('$($group.Id)')"

            MemberSettings = @{
                AllowCreateUpdateChannels           = $false   # Add and edit channels OFF
                AllowCreatePrivateChannels          = $false   # Add private channels OFF
                AllowUpdatePrivateChannels          = $false   # Edit private channels OFF
                AllowDeleteChannels                 = $false   # Delete channels OFF
                AllowAddRemoveApps                  = $false   # Add/edit/remove apps OFF
                AllowCreateUpdateRemoveTabs         = $false   # Add/edit/remove tabs OFF
                AllowCreateUpdateRemoveConnectors   = $false   # Add/edit/remove connectors OFF
                AllowEditOwnMessages                = $true    # Edit sent messages ON
                AllowDeleteOwnMessages              = $true    # Delete sent messages ON
                AllowDeleteMessages                 = $true    # Owners can delete messages ON
            }

            GuestSettings = @{
                AllowCreateUpdateChannels = $false   # Guests can add/edit channels OFF
                AllowDeleteChannels       = $false   # Guests can delete channels OFF
            }

            MessagingSettings = @{
                AllowUserEditMessages    = $true     # Edit sent messages ON
                AllowUserDeleteMessages  = $true     # Delete sent messages ON
                AllowOwnerDeleteMessages = $true     # Owners can delete messages ON
                AllowTeamMentions        = $true     # Mention teams ON
                AllowChannelMentions     = $true     # Mention channels ON
            }

            FunSettings = @{
                AllowGiphy            = $true        # Giphy ON
                GiphyContentRating    = "Moderate"   # PG ≈ Moderate
                AllowStickersAndMemes = $true        # Stickers and memes ON
                AllowCustomMemes      = $true        # Custom memes ON
            }
        }

        $team = New-MgTeam -BodyParameter $teamParams
        Write-Host "New-MgTeam returned without throwing for group '$($group.DisplayName)'"
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Host "New-MgTeam error: $errMsg"

        try {
            Start-Sleep -Seconds 5
            $checkTeam = Get-MgTeam -TeamId $group.Id -ErrorAction SilentlyContinue
            if ($checkTeam) {
                Write-Host "Team appears to exist despite New-MgTeam error for group '$($group.DisplayName)'"
                $errMsg = "Team seems created, but New-MgTeam reported error: $errMsg"
            }
        }
        catch { }

        Write-ErrorCsv -Owner $ownerUpn `
                       -MailNickName $mailNickname `
                       -DisplayName $displayName `
                       -SourceM365Group $sourceMail `
                       -ErrorMessage $errMsg
        continue
    }

    Write-Host "Team enabled for group '$($group.DisplayName)' with custom settings"
}
