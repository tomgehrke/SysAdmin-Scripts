<#
Watch-PingTimes.ps1
#>

Param( 
    [Parameter(Mandatory=$true)][string]$Target,
    [Parameter(Mandatory=$false)][string]$FileName,
    [Parameter(Mandatory=$false)][switch]$Notify,
    [Parameter(Mandatory=$false)][int]$Interval=5,
    [Parameter(Mandatory=$false)][int]$NotifyThreshold=10,
    [Parameter(Mandatory=$false)][switch]$ShowOutput
)

# Load the relevant .NET assembly
Add-Type -AssemblyName System.Windows.Forms

if ($Notify) {
    $NotificationIcon = New-Object System.Windows.Forms.NotifyIcon
    $ProcessPath = (Get-Process -id $pid).Path
    $NotificationIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ProcessPath) 
    $NotificationIcon.BalloonTipTitle = "Ping Alert!" 
}

if (!$FileName) {
    $FileName = "pingtimes.csv"
}

$Source = (Get-NetIPAddress -AddressFamily IPv4 -InterfaceAlias Ethernet).IPAddress

Write-Output "Watching pings to ""$Target""..."

while ($Continuous -or $Counter -le $Number) {
    $ErrorMessage = $null
    $ResponseTime = $null
    $Now = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")

    try {
        $ResponseTime = (Get-WmiObject Win32_PingStatus -Filter "Address='192.168.141.56'").ResponseTime
    } catch {
        $ErrorMessage = $_.Exception.Message + $_
    }
 
    $Result = [PSCustomObject]@{
        "Timestamp" = $Now
        "Source" = $Source
        "Target" = $Target
        "ResponseTime" = $ResponseTime
        "ErrorMessage" = $ErrorMessage
        }

    if ($FileName) {
        $Result | Export-Csv -Path $FileName -Append -NoTypeInformation
    }

    if ($ShowOutput) {
        clear
        Write-Output $Result | Format-Table -Autosize
    }

    if ($Notify -and $ResponseTime -ge $NotifyThreshold) {
        $NotificationIcon.BalloonTipText = "Ping response time of ""$ResponseTime"" is greater than threshold of ""$NotifyThreshold"""
        $NotificationIcon.Visible = $true 
        $NotificationIcon.ShowBalloonTip(5000)
    }

    Start-Sleep $Interval
}
