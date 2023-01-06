# Backup-FolderAcls.ps1
#
# Arguments
# ===================
# -Path         | The root path to scan
# -BackupFolder | The path to the folder where backup files will be created
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
    [switch]
    $PermissionNotInherited
    )

Add-Type -AssemblyName System.Windows.Forms

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

if (!(Test-Path $Path -PathType Container)) {
    Write-Output "[ERROR] ""$Path"" does not exist!"
    Exit
}

if ($Path.Length -gt 3) {
    $Path = $Path.TrimEnd("\")
}

function Backup-Acl {
    Param( 
        [Parameter(Mandatory=$true)]
        [string]
        $Path
        ,
        [Parameter(Mandatory=$true)]
        [string]
        $Output
        )

    $Backup = @()

    try {
        $Acl = Get-Acl $Path
        $Shares = Get-SmbShare | Where-Object {$_.Path -eq $Path} | Select-Object -ExpandProperty Name
        foreach ($Permission in $Acl.Access) {
            foreach ($Share in $Shares) {
                if (!$PermissionNotInherited -or ($PermissionNotInherited -and !$Permission.IsInherited)) {
                    $Backup += [PSCustomObject]@{
                        "Path" = $Path
                        "Share" = $Share
                        "FileSystemRights" = $Permission.FileSystemRights
                        "AccessControlType" = $Permission.AccessControlType
                        "IdentityReference" = $Permission.IdentityReference
                        "IsInherited" = $Permission.IsInherited
                        "InheritanceFlags" = $Permission.InheritanceFlags
                        "PropagationFlags" = $Permission.PropagationFlags
                        "Exception" = ""
                        } 
                    }
                }
            }
    } catch {
        $Backup += [PSCustomObject]@{
            "Path" = $Path
            "Share" = ""
            "FileSystemRights" = ""
            "AccessControlType" = ""
            "IdentityReference" = ""
            "IsInherited" = ""
            "InheritanceFlags" = ""
            "PropagationFlags" = ""
            "Exception" = $Error[0].Exception.Message
            }
    }

    $Backup | Export-Csv -Path $Output -NoTypeInformation -Append
}

# Check for backup folder if provided and create if it doesn't exist
if ($BackupFolder -ne "" -and !(Test-Path -Path $BackupFolder)) {
    Write-Host "- Backup folder does not exist. Creating it..."
    $null = New-Item -Path $BackupFolder -ItemType Directory -Force
}

$BackupFileName = $BackupFile
if ($BackupFileName -eq "") {
    $Now = Get-Date -Format "yyyyMMddHHmm"
    $BackupFileName = "FolderAclBackup-$Now.csv"
}

Backup-Acl -Path $Path -Output "$BackupFolder\$BackupFileName"

Write-Output "- Loading directory tree..."
$Folders = Get-ChildItem $Path -Directory -Recurse -ErrorAction SilentlyContinue
$FolderCount = $Folders.Count
$CurrentFolder = 0

foreach ($Folder in $Folders) {
    $CurrentFolder++
    Write-Progress -Id 0 -Activity "Scanning Folders" -PercentComplete ($CurrentFolder/$FolderCount*100) -Status "Scanning $Folder"
    Backup-Acl -Path $Folder.FullName -Output "$BackupFolder\$BackupFileName"
}

Write-Output "- Backup created - ""$BackupFolder\$BackupFileName"""
