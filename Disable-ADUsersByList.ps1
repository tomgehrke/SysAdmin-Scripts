# Disable-ADUsersByList
#
# Intended to facilitate the disabling of users in bulk with minimal interaction.
#
# If a user is identified in the listing by somethin unique like sAMAccountName or
# distinguishedName, they will be easily found and disabled.
#
# If a user is identified by name, the script will attempt to find them via several
# methods going from specific to general. Because this is not exact, the operator will
# always be presented with a list of users to confirm even if only one object is returned.
#
# A report detailing the disposition of each user on the list will be generated as a CSV.
#
# Arguments
# ===================
# -Server      | Target domain [REQUIRED]
# -UserList    | User list
# -ReportPath  | Output filename (defaults to "disable-user-result-<DATETIME>.csv")
# -Description | String to place in disabled user's Description
# -TestMode    | Generates report based on what it would have done, but does not actually disable any users

Param( 
    [Parameter(Mandatory=$false)]
    [string]
    $Server=""
    , 
    [Parameter(Mandatory=$false)]
    [string]
    $UserList=""
    ,
    [Parameter(Mandatory=$false)]
    [string]
    $ReportPath="disable-user-result-$(Get-Date -Format ""yyyyMMddHHmm"").csv"
    ,
    [Parameter(Mandatory=$false)]
    [string]
    $Description=""
    ,
    [Parameter(Mandatory=$false)]
    [switch]
    $TestMode
    )

Add-Type -AssemblyName System.Windows.Forms

# User didn't pass in a list, ask them for one.
if ($UserList -eq "") {
    $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $FileDialog.Filter = "User list (*.txt)|*.txt"
    $ShowForm = $FileDialog.ShowDialog()
    if ($ShowForm -eq "OK") {
        $UserList = $FileDialog.FileName
    } else {
        Write-Output "[ERROR!] A list of users is required!"
        Exit
    }
}

# User passed a list, but is it valid?
if (!(Test-Path $UserList -PathType Leaf)) {
    Write-Output "[ERROR] ""$UserList"" does not exist!"
    Exit
}

$Users = Get-Content $UserList
$TotalUsers = $Users.Length
$CurrentUser = 0
$Results = @()

foreach ($User in $Users) {
    
    # Show progress
    $CurrentUser++
    Write-Progress -Id 0 -Activity "User Search" -PercentComplete ($CurrentUser/$TotalUsers*100) -Status "- Searching for $User..."

    # Prepare for possibility of additional switches
    $AdditionalSwitches = @{}
    
    $UserName = $User
    $Domain = $Server
    $FoundByID = $false

    # Check for domain in user@domain format
    if ($User.Contains("@")) {
        $UserParts = $User -split "@"
        $UserName = $UserParts[0]
        $Domain = $UserParts[1]
    # Check for domain in distinguishedName
    } elseif ($User.Contains("DC=")) {
        $Domain = $User -replace "DC=([^,]+(?:,|$))|.", '$1' -replace ",", "."
    # Check for domain in user@domain format
    } elseif ($User.Contains("\")) {
        $UserParts = $User -split "\\"
        $UserName = $UserParts[1]
        $Domain = $UserParts[0]
    }

    # Override domain if one was explicitly passed
    if ($Server) {
        $Domain = $Server
    }

    # Do we have a target Domain?
    if ($Domain -ne "") {
        $AdditionalSwitches.Add("Server", $Domain)
    }

    Write-Output "[!] Searching for ""$UserName"" $(if($Domain){""in $Domain""}) "

    # Look for user
    try {
        $FoundUser = Get-ADUser -Identity $UserName @AdditionalSwitches
        $FoundByID = $true
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        $FoundUser = $null
    } catch [Microsoft.ActiveDirectory.Management.ADServerDownException] {
        $Results += [PSCustomObject]@{
            "User" = $User
            "Result" = "Server/Domain unavailable"
            }
        Continue
    }

    # User not found yet, try display name
    if (!($FoundUser)) {
        try {
            $FoundUser = Get-ADUser -Filter {displayName -eq $UserName} @AdditionalSwitches
        } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $FoundUser = $null
        }
    }

    # User not found yet, try matching whatever was passed with the last name
    if (!($FoundUser)) {
        try {
            $FoundUser = Get-ADUser -Filter {surname -eq $UserName} @AdditionalSwitches
        } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $FoundUser = $null
        }
    }
    
    # Need to get more granular at this point and do some name parsing
    if (!($FoundUser)) {
        $LastName = ""
        $FirstName = ""

        # Assume lastname, firstname if there's a comma
        if ($UserName.Contains(",")) {
            $NameParts = $UserName -split ","
            $LastName = $NameParts[0]
            if ($NameParts[1].Trim().Contains(" ")) {
                $FirstName = $($NameParts[1].Trim() -split " ")[0]
            } else {
                $FirstName = $NameParts[1].Trim()
            }
        # Assume firstname [middle name] lastname if there are spaces
        } elseif ($UserName.Contains(" ")) {
            $NameParts = $UserName -split " "
            $LastName = $NameParts[-1]
            $FirstName = $NameParts[0]
        }

        Write-Verbose "First Name: $FirstName"
        Write-Verbose "Last Name: $LastName"
    }
    
    # User not found yet, try last name and first name
    if (!($FoundUser) -and $LastName -ne "" -and $FirstName -ne "") {
        try {
            Write-Verbose "Searching by lastname, firstname..."
            $FoundUser = Get-ADUser -Filter {surname -eq $LastName -and givenName -eq $FirstName} @AdditionalSwitches
        } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $FoundUser = $null
        }
    }

    # User not found yet, try last name and try to match a shortened first name
    if (!($FoundUser) -and $LastName -ne "" -and $FirstName -ne "") {
        try {
            Write-Verbose "Searching by lastname and a shortened first name..."
            $FoundUser = Get-ADUser -Filter "surname -eq '$LastName' -and givenName -like '$FirstName*'" @AdditionalSwitches
        } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $FoundUser = $null
        }
    }

    # User not found yet, try just the last name
    if (!($FoundUser) -and $LastName -ne "") {
        try {
            Write-Verbose "Searching by lastname..."
            $FoundUser = Get-ADUser -Filter {surname -eq $LastName} @AdditionalSwitches
        } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            $FoundUser = $null
        }
    }

    # All search options exhausted. Giving up!
    if (!($FoundUser)) {
        Write-Output "    NOT FOUND!"
        $Results += [PSCustomObject]@{
            "User" = $User
            "Result" = "NOT FOUND!"
            }
        Continue
    }

    # Not found by identity, so any users found at this point need to be reviewed
    if (!$FoundByID) {
        $SelectedUser = $FoundUser | Out-GridView -Title "Select the correct object for ""$User""" -PassThru
        if ($SelectedUser) {
            $FoundUser = $SelectedUser
        } else {
            Write-Output "    NONE SELECTED!"
            $Results += [PSCustomObject]@{
                "User" = $User
                "Result" = "NONE SELECTED!"
                }
            Continue
        }
    }

    if ($FoundUser.Enabled) {
        Set-ADUser -Identity $FoundUser -Description $Description -Enabled $False -WhatIf:$TestMode
        Write-Output "    Disabled!"
        $Results += [PSCustomObject]@{
            "User" = $User
            "Result" = "User disabled"
            "sAMAccountName" = $FoundUser.sAMAccountName
            "distinguishedName" = $FoundUser.distinguishedName
            }
    } else {
        Write-Output "    Nothing done. Previously disabled."
        $Results += [PSCustomObject]@{
            "User" = $User
            "Result" = "User was previously disabled"
            "sAMAccountName" = $FoundUser.sAMAccountName
            "distinguishedName" = $FoundUser.distinguishedName
            }
    }

}

$Results | Export-Csv $ReportPath -NoTypeInformation