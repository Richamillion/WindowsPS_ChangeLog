$machine = hostname
$dateTime = Get-Date -Format "yyyyMMdd_HHmmss"
$changeLogsFolder = "SOMEPATH"
Get-WmiObject -Class Win32_Product | Select-Object -Property Name, Version | Sort-Object Name | Export-Csv "$changeLogsFolder\$machine.startup.$dateTime.csv" -NoTypeInformation
