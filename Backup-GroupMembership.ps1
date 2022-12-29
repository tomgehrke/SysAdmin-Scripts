# Remove group members with backup
#
# Arguments
# ===================
# -Server        | Target domain [REQUIRED]
# -FileName      | The text file containing a list of groups (either CN or DN) to process
# -BackupFolder  | Target location for group backup (gbak) files
# -RemoveMembers | Switch to subsequently removes the group's members after backing them up

Param( 
    [Parameter(Mandatory=$true)][string]$Server, 
    [Parameter(Mandatory=$false)][string]$FileName="",
    [Parameter(Mandatory=$false)][string]$BackupFolder=".",
    [Parameter(Mandatory=$false)][switch]$RemoveMembers,
    [Parameter(Mandatory=$false)][switch]$Recursive
    )

Add-Type -AssemblyName System.Windows.Forms

if ($FileName -eq "") {
    $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $ShowForm = $FileDialog.ShowDialog()
    if ($ShowForm -eq "OK") {
        $FileName = $FileDialog.FileName
    } else {
        Write-Host "[ERROR!] A list of groups in a text file must be provided!"
        Exit
    }
}

if (!(Test-Path $FileName -PathType Leaf)) {
    Write-Host "[ERROR] ""$FileName"" does not exist!"
    Exit
}

$GroupList = Get-Content $FileName
$TotalGroups = $GroupList.Count
$Now = Get-Date -Format "yyyyMMddHHmm"

# Check for backup folder if provided and create if it doesn't exist
if ($BackupFolder -ne "" -and !(Test-Path -Path $BackupFolder)) {
    Write-Host "Backup folder does not exist. Creating it..."
    $null = New-Item -Path $BackupFolder -ItemType Directory -Force
}

$CurrentGroup = 0

# Go through the list of groups
foreach ($Group in $GroupList) {
    $CurrentGroup++;
    Write-Progress -Id 0 -Activity "Scanning Groups" -PercentComplete ($CurrentGroup/$TotalGroups*100) -Status "Scanning $Group"

    # Create backup filename by removing any non-word characters from the group name
    try {
        $ThisGroup = Get-ADGroup -Identity $Group -Server $Server -Properties Name, distinguishedName, GroupCategory, GroupScope, Description
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Host "[ERROR] $Group was not found!"
        Continue
    }

    $BackupFileName = $ThisGroup.Name -Replace ('\W','')
    
    $GroupMembers = Get-ADGroupMember -Server $Server -Identity $Group -Recursive:$Recursive | Select-Object -ExpandProperty distinguishedName
    $TotalMembers = $GroupMembers.Count

    if ($TotalMembers -gt 0) {    
        
        # Create the backup
        Set-Content -Path "$BackupFolder\$BackupFileName-$Now.gbak" -Value "[GROUP]$($ThisGroup.Name)"
        Add-Content -Path "$BackupFolder\$BackupFileName-$Now.gbak" -Value "[CATEGORY]$($ThisGroup.GroupCategory)"
        Add-Content -Path "$BackupFolder\$BackupFileName-$Now.gbak" -Value "[SCOPE]$($ThisGroup.GroupScope)"
        Add-Content -Path "$BackupFolder\$BackupFileName-$Now.gbak" -Value "[DESCRIPTION]$($ThisGroup.Description)"
        Add-Content -Path "$BackupFolder\$BackupFileName-$Now.gbak" -Value "[SID]$($ThisGroup.SID)"
        Add-Content -Path "$BackupFolder\$BackupFileName-$Now.gbak" -Value $GroupMembers

        if ($RemoveMembers) {
            $CurrentMember = 0
            Write-Progress -Id 1 -ParentId 0 -Activity "Removing Members" -PercentComplete 0

            # Remove members
            foreach ($Member in $GroupMembers) {
                $CurrentMember++;

                # Members may come from different domains. To avoid "referral" errors, get the user object from its domain.
                # That means pulling its domain out of its distinguishedName.
                $MemberDomain = $Member.Substring($Member.IndexOf("DC=")+3).Replace(",DC=",".")
                $ThisUser = Get-ADUser -Server $MemberDomain -Identity $Member

                Write-Progress -Id 1 -ParentId 0 -Activity "Removing Members" -PercentComplete ($CurrentMember/$TotalMembers*100) -Status "Removing $Member"
                Remove-ADGroupMember -Server $Server -Identity $Group -Members $ThisUser -Confirm:$false
            }

            # Clean up any remaining subgroups
            Write-Progress -Id 1 -ParentId 0 -Activity "Removing Members" -Status "Cleaning up sub-groups"
            $SubGroups = Get-ADGroupMember -Server $Server -Identity $Group | Select-Object -ExpandProperty distinguishedName
            foreach ($SubGroup in $SubGroups) {
                $MemberDomain = $SubGroup.Substring($SubGroup.IndexOf("DC=")+3).Replace(",DC=",".")
                $ThisSubGroup = Get-ADGroup -Server $MemberDomain -Identity $SubGroup
                Remove-ADGroupMember -Server $Server -Identity $Group -Members $ThisSubGroup -Confirm:$false
            }
        }
    }    
}