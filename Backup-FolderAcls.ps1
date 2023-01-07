# Backup-FolderAcls.ps1
#
# Arguments
# ===================
# -Path                   | The root path to scan
# -BackupFile             | The path to the folder where backup files will be created
# -BackupFolder           | The path to the folder where backup files will be created
# -Server                 | Does nothing but populate the "Server" column in the output file
# -PermissionNotInherited | Only list folders where permissions are explicitly set and not inherited
#
# NOTE
# ===================
# If you receive errors about paths not being found, it is possible that you
# are exceeding the maximum of 260 characters for a path in the Windows API.
#
# You may, however, extend this length to ~32,767 characters by prefixing the
# path in one of two ways.
#
# 1. Paths with Drive mappings
#
# \\?\C:\<very long path>
#
# 2. UNC paths
#
# \\?\UNC\<server>\<share>

Param( 
    [Parameter(Mandatory=$false)]
    [string]
    $Path=""
    ,
    [Parameter(Mandatory=$false)]
    [string]
    $BackupFile=""
    ,
    [Parameter(Mandatory=$false)]
    [string]
    $BackupFolder="."
    ,
    [Parameter(Mandatory=$false)]
    [string]
    $Server=""
    ,
    [Parameter(Mandatory=$false)]
    [switch]
    $PermissionNotInherited
    )

Add-Type -AssemblyName System.Windows.Forms

# A path was not provided on the command line. Let the user select one.
if ($Path -eq "") {
    $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderDialog.Description = "Select the root folder"
    $FolderDialog.RootFolder = "MyComputer"
    $FolderDialog.SelectedPath = $(Get-Location).Path
    $FolderDialog.ShowNewFolderButton = $false
    $ShowForm = $FolderDialog.ShowDialog()
    if ($ShowForm -eq "OK") {
        $Path = $FolderDialog.SelectedPath
    } else {
        Write-Output "[ERROR!] A target path must be selected!"
        Exit
    }
}

# Make sure the path exists. Exit if it does not.
if (!(Test-Path $Path -PathType Container)) {
    Write-Output "[ERROR] ""$Path"" does not exist!"
    Exit
}

# If the path is longer than 3 characters, We're probably referencing
# more than just the root of a drive so we want to trim that last backslash.
# If the path is 3 characters (e.g.; "D:\"), we want to keep that
# trailing backslash.
if ($Path.Length -gt 3) {
    $Path = $Path.TrimEnd("\")
}

# Check for backup folder if provided and create it if it doesn't exist
if ($BackupFolder -ne "" -and !(Test-Path -Path $BackupFolder)) {
    Write-Output "- Backup folder does not exist. Creating it..."
    $null = New-Item -Path $BackupFolder -ItemType Directory -Force
}

# Set the backup filename to whatever was passed
$BackupFileName = $BackupFile

# If a backup filename was not passed, set a default value
if ($BackupFileName -eq "") {
    $Now = Get-Date -Format "yyyyMMddHHmm"
    $BackupFileName = "FolderAclBackup-$Now.csv"
}

# This function is responsible for outputting the ACL and Share information
# for a folder including any errors that occur (most likely "access denied").
function Backup-Acl {
    Param( 
        [Parameter(Mandatory=$true)]
        [string]
        $FolderPath
        ,
        [Parameter(Mandatory=$true)]
        [string]
        $Output
        )

    # Initialize the backup
    $Backup = @()
    
    # Get the ACL for the folder. Capture the error if it fails.
    try {
        $Acl = Get-Acl $FolderPath
    } catch {
        $AclError = $Error[0].Exception.Message
    }

    # Get any Share information for the folder. Capture the error if it fails.
    # Note that most folders will probably not be shared and some folders may be
    # shared multiple times.
    try {
        $Shares = Get-SmbShare | Where-Object {$_.Path -eq $FolderPath} | Select-Object -ExpandProperty Name
    } catch {
        $ShareError = $Error[0].Exception.Message
    }

    # If there was no Share information, we still need a single item so
    # we'll create an empty one.
    if (!$Shares) {
        $Shares = @("")
    }

    # If reading the ACL resulted in an error, we still need an item so
    # we'll create an empty one.
    if ($AclError) {
        $Owner = ""
        $Permissions = [PSCustomObject]@{
            "FileSystemRights" = ""
            "AccessControlType" = ""
            "IdentityReference" = ""
            "IsInherited" = ""
            "InheritanceFlags" = ""
            "PropagationFlags" = ""
            }
    } else {
        # We got an ACL so we capture the attributes we'll be backing up.
        $Permissions = $Acl.Access
        $Owner = $Acl.Owner
    }

    # Loop through all of the permissions assigned in the ACL
    foreach ($Permission in $Permissions) {
        # If the user only wanted to see folders where permissions are explicitly assigned (not inherited),
        # we check the IsInherited flag and skip the rest of the loop if it is TRUE. But only if
        # no errors were reported from pulling the ACL or Share information. We always report those no
        # matter what.
        if ($PermissionNotInherited -and $Permission.IsInherited -and !$AclError -and !$ShareError) {
          continue   
        }

        # Loop through all of the Shares. 
        # If the folder was not shared, the share information will be the blank item we created.
        # If the folder was shared multiple times, the report will reflect repeated information for 
        # the ACL (because that isn't Share-specific) with the only difference being the Share name.
        foreach ($Share in $Shares) {
            $Backup += [PSCustomObject]@{
                "Server" = $Server
                "Path" = $FolderPath
                "Share" = $Share
                "Owner" = $Owner
                "FileSystemRights" = $Permission.FileSystemRights
                "AccessControlType" = $Permission.AccessControlType
                "IdentityReference" = $Permission.IdentityReference
                "IsInherited" = $Permission.IsInherited
                "InheritanceFlags" = $Permission.InheritanceFlags
                "PropagationFlags" = $Permission.PropagationFlags
                "Exceptions" = "$AclError`n$ShareError"
                } 
        }

    }

    # Write the folder's information to our backup file
    $Backup | Export-Csv -Path $Output -NoTypeInformation -Append
}

# Write the root path information to our backup file
Backup-Acl -FolderPath $Path -Output "$BackupFolder\$BackupFileName"

# Recursively populate the collection of folders in the target folder's tree
Write-Output "- Loading directory tree..."
$Folders = Get-ChildItem $Path -Directory -Recurse -ErrorAction SilentlyContinue
$FolderCount = $Folders.Count
$CurrentFolder = 0

# Loop through all of the folders calling our Backup-Acl function for each one
foreach ($Folder in $Folders) {
    $CurrentFolder++
    Write-Progress -Id 0 -Activity "Scanning Folders" -PercentComplete ($CurrentFolder/$FolderCount*100) -Status "Scanning $($Folder.FullName)"
    Backup-Acl -FolderPath $Folder.FullName -Output "$BackupFolder\$BackupFileName"
}

# We're done!
Write-Output "- Backup created - ""$BackupFolder\$BackupFileName"""