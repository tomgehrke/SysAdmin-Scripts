Get-NetAdapterBinding -ComponentID ms_tcpip6 | Where-Object {$_.Enabled -eq $True} | ForEach-Object {Set-NetAdapterBinding -Name $_.Name -ComponentID $_.ComponentID -Enabled $False}
Write-Output "Final Status..."
Get-NetAdapterBinding -ComponentID ms_tcpip6