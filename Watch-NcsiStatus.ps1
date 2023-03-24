<# 
Watch-NcsiStatus.ps1

Arguments
===================
-Continuous | Run until interrupted
-FileName   | The path/filename of the report
-Interval   | Interval in seconds
-Notify     | Notify on status change
-Number     | Number of checks to run
#>

Param( 
    [Parameter(Mandatory=$false)][switch]$Continuous,
    [Parameter(Mandatory=$false)][string]$FileName="",
    [Parameter(Mandatory=$false)][int]$Interval=60, 
    [Parameter(Mandatory=$false)][switch]$Notify,
    [Parameter(Mandatory=$false)][int]$Number=5 
    )

# Load the relevant .NET assembly
Add-Type -AssemblyName System.Windows.Forms

if ($Notify) {
    $NotificationIcon = New-Object System.Windows.Forms.NotifyIcon
    $ProcessPath = (Get-Process -id $pid).Path
    $NotificationIcon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($ProcessPath) 
    $NotificationIcon.BalloonTipTitle = "Network Change!" 
}

$Counter = 1
$CurrentNetworkStatus = ""
$CurrentInternetStatus = ""

while ($Continuous -or $Counter -le $Number) {
    if (!$Continuous) {$Counter++}

    $ErrorMessage = ""
    $HttpResponse = ""
    $NetworkStatus = if ([System.Windows.Forms.SystemInformation]::Network -eq $true) {"Yes"} else {"No"}
    try {
        # $HttpResponse = (Invoke-WebRequest -Uri "http://www.msftncsi.com/ncsi.txt" -DisableKeepAlive).Content
        $HttpResponse = (Invoke-WebRequest -Uri "http://www.msftconnecttest.com/connecttest.txt" -DisableKeepAlive).Content
    } catch {
        $ErrorMessage = $_.Exception.Message + $_
    }
    # $InternetStatus = if ($HttpResponse -eq "Microsoft NCSI") {"Yes"} else {"No"}
    $InternetStatus = if ($HttpResponse -eq "Microsoft Connect Test") {"Yes"} else {"No"}

    $Result = [PSCustomObject]@{
        "Timestamp" = Get-Date
        "Network" = $NetworkStatus
        "Internet" = $InternetStatus
        "Error" = $ErrorMessage
        }

    Write-Output $Result

    if ($Notify -and ($NetworkStatus -ne $CurrentNetworkStatus -or $InternetStatus -ne $CurrentInternetStatus)) {
        $NotificationIcon.BalloonTipText = "Network Status is ""$NetworkStatus"" and Internet Status is ""$InternetStatus"""
        $NotificationIcon.Visible = $true 
        $NotificationIcon.ShowBalloonTip(5000)
        $CurrentNetworkStatus = $NetworkStatus
        $CurrentInternetStatus = $InternetStatus
    }

    if ($FileName) {
        $Result | Export-Csv -Path $FileName -Append -NoTypeInformation
    }

    if ($Continuous -or $Counter -le $Number) {
        Start-Sleep $Interval
    }
}

