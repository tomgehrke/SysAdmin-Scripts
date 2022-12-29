# Restore group members from backup
#
# Arguments
# ===================
# -Server       | Target domain [REQUIRED]
# -FileName     | Group backup file
# -CreateGroup  | Create a new group using backed up attribues if it does not exist
# -NewName      | Explicitly set the new group's Name
# -NewScope     | Explicitly set the new group's GroupScope
# -NewCategory  | Explicitly set the new group's GroupCategory
# -TestMode     | Do not actually add users

Param( 
    [Parameter(Mandatory=$true)][string]$Server, 
    [Parameter(Mandatory=$false)][string]$FileName="",
    [Parameter(Mandatory=$false)][switch]$CreateGroup,
    [Parameter(Mandatory=$false)][string]$NewName="",
    [Parameter(Mandatory=$false)][string]$NewScope="",
    [Parameter(Mandatory=$false)][string]$NewCategory="",
    [Parameter(Mandatory=$false)][switch]$TestMode
    )

Add-Type -AssemblyName System.Windows.Forms

if ($FileName -eq "") {
    $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $FileDialog.Filter = "Group Backup (*.gbak)|*.gbak"
    $ShowForm = $FileDialog.ShowDialog()
    if ($ShowForm -eq "OK") {
        $FileName = $FileDialog.FileName
    } else {
        Write-Output "[ERROR!] A group backup file must be provided!"
        Exit
    }
}

if (!(Test-Path $FileName -PathType Leaf)) {
    Write-Output "[ERROR] ""$FileName"" does not exist!"
    Exit
}

try {
    $MaxHeaderRows = 5
    $ActualHeaderRows = 0
    
    # Pull original group information from backup file header
    $Header = (Get-Content $FileName | Select -First $MaxHeaderRows)

    for ($i=0; $i -le $MaxHeaderRows - 1; $i++) {
        $HeaderItem = $Header[$i].Split("]")

        switch ($HeaderItem[0]) {
            "[GROUP" {
                $GroupName = $HeaderItem[1]
                $ActualHeaderRows++
                break
                }
            "[CATEGORY" {
                $GroupCategory = $HeaderItem[1]
                $ActualHeaderRows++
                break
                }
            "[SCOPE" {
                $GroupScope = $HeaderItem[1]
                $ActualHeaderRows++
                break
                }
            "[DESCRIPTION" {
                $Description = $HeaderItem[1]
                $ActualHeaderRows++
                break
                }
            "[SID" {
                $SID = $HeaderItem[1]
                $ActualHeaderRows++
                break
                }
        }
    }
} catch {
    Write-Output "[ERROR] Problem reading backup file header"
    Exit
}

$GroupMembers = (Get-Content $FileName | Select -Skip $ActualHeaderRows)
$TotalMembers = $GroupMembers.Count

if ($TotalMembers -gt 0) {

    # Override values with explicit parameters
    if ($NewName -ne "") {
        $GroupName = $NewName
    }

    Write-Output "Beginning restore of ""$GroupName"" on ""$Server""..."

    try {
        # Get the target group
        $Group = Get-ADGroup -Identity "$GroupName" -Server "$Server"
        Write-Output '- Group found...'
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        # The target group does not exist
        if ($CreateGroup -or ($Host.UI.PromptForChoice("Group does not exist!", "Do you want to create the group?", $("&Yes","&No"), 1) -eq 0)) {

            # Override values with explicit parameters
            if ($NewCategory -ne "") {
                $GroupCategory = $NewCategory
            }
            if ($NewScope -ne "") {
                $GroupScope = $NewScope
            }

            # Create the group
            Write-Output "- Creating group..."
            New-ADGroup -Server "$Server" -Name "$GroupName" -GroupCategory "$GroupCategory" -GroupScope "$GroupScope" -Description "$Description"
            $Group = Get-ADGroup "$GroupName" -Server "$Server" 
        } else {
            Write-Output "[ERROR] Group does not exist!"
            Exit
        }
    }

    # Do we have a group object at this point
    if ($Group -ne $null) {

        $CurrentMember = 0

        foreach ($Member in $GroupMembers) {
            $CurrentMember++;
            Write-Progress -Id 0 -Activity "Adding Members" -PercentComplete ($CurrentMember/$TotalMembers*100) -Status "- Adding $($ThisUser.Name)..."

            # Members may come from different domains. To avoid "referral" errors, get the user object from its domain.
            # That means pulling its domain out of its distinguishedName.
            $MemberDomain = $Member.Substring($Member.IndexOf("DC=")+3).Replace(",DC=",".")
            $ThisUser = Get-ADUser -Server $MemberDomain -Identity $Member

            Add-ADGroupMember -Identity $Group -Members $ThisUser -Confirm:$false -WhatIf:$TestMode
        }
    }
} else {
    Write-Output "[ERROR] No group members found in backup file!"
}