$machine = hostname
$dateTime = Get-Date -Format "yyyyMMdd_HHmmss"

### Change Log Shutdown Export
Write-Host "Updating software changelog..."

# Path to the Change Log Folder
$changeLogsFolder = "SOMEPATH"

# Get the Change Log from the Unlock_Update Script and filter it to ensure it is a .CSV file
$changeLogs = Get-ChildItem -Path $changeLogsFolder -Filter "*.csv" | Where-Object { $_.Name -like "$machine.startup*" } | Sort-Object Name -Descending | Select-Object -First 1

# Check if changelog export from unlock_update exists or notand if so, compare and export changes
if ($null -eq $changeLogs) {
    
    Write-Host "No current startup change log found. Checking recently installed software and event log for changes..."
    
    $currentChangeLog = Get-WmiObject -Class Win32_Product | Select-Object -Property Name, Version, InstallDate | Sort-Object Name
    $checkDate = Get-Date
    # How many days to check?
    $withinDays = $checkDate.AddDays(-3)

    # Check for recent changes based on how many days
    $recentChanges = $currentChangeLog | Where-Object { $_.InstallDate -ne $null -and [datetime]::ParseExact($_.InstallDate, 'yyyyMMdd', $null) -gt $withinDays }

    # Prepare to hold recent changes
    $installedSoftware = @()

    foreach ($install in $recentChanges) {
        # Parse the software name and version from recent changes
        $installedSoftware += [PSCustomObject]@{
            Name        = $install.Name
            FromVersion = "N/A"
            ToVersion   = $install.Version
            Status      = "Installed"
        }
    }

    # Check the Event Log for recently uninstalled software
    $uninstallEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'Application'; 
        ID = 1034
        StartTime = $withinDays
    } | Select-Object Message

    # Prepare to hold uninstalled software
    $uninstalledSoftware = @()

    foreach ($uninstall in $uninstallEvents) {
        # Parse the software name and version from the event log
        if ($uninstall.Message -match 'removed the product\. Product Name: (.+?)\. Product Version: (.+?)\. Product Language') {
            $softwareName = $matches[1].Trim()
            $softwareVersion = $matches[2].Trim()
            
            $uninstalledSoftware += [PSCustomObject]@{
                Name = $softwareName
                FromVersion = $SoftwareVersion
                ToVersion = "N/A"
                Status = "Removed"
            }
        }
    }

    # Combine recent installations and removals
    $softwareEvents = $installedSoftware + $uninstalledSoftware

    # Export to CSV if there are changes
    if ($softwareEvents.Count -gt 0) {
        $softwareEvents | Export-Csv "$changeLogsFolder\$machine.changelog.$dateTime.csv" -NoTypeInformation
        
        Write-Host "Changes written to change log."
    } else {
        
        Write-Host "No recent software changes."
    }

} else {
    # Import the Last Software Change Log for Comparison
    $lastChangeLog = Import-Csv $changeLogs.FullName

    # Get the Current Software Change Log
    $currentChangeLog = Get-WmiObject -Class Win32_Product | Select-Object -Property Name, Version | Sort-Object Name
    Write-Host " done!" -NoNewline

    Write-Host "Exporting changes to changelog..."
    # Create a table of the Last Change Log for Comparison
    $lastChange = @{}
    foreach ($lastSoftware in $lastChangeLog) {
        if ($lastSoftware.Name) {  # Ensure Name is not null or empty
            $lastChange[$lastSoftware.Name] = $lastSoftware.Version
        }
    }

    # Array to hold Software Changes
    $softwareChanges = @()
    foreach ($currentSoftware in $currentChangeLog) {
        if ($currentSoftware.Name) {  # Ensure Name is not null or empty
            if ($lastChange.ContainsKey($currentSoftware.Name)) {
                if ($lastChange[$currentSoftware.Name] -ne $currentSoftware.Version) {
                    $softwareChanges += [PSCustomObject]@{
                        Name            = $currentSoftware.Name
                        FromVersion     = $lastChange[$currentSoftware.Name]
                        ToVersion       = $currentSoftware.Version
                        Status          = "Updated"
                    }
                }
            } else {
                $softwareChanges += [PSCustomObject]@{
                    Name            = $currentSoftware.Name
                    FromVersion     = 'Not Found'
                    ToVersion       = $currentSoftware.Version
                    Status          = "New Install"
                }
            }
        }
    }

    # Check for any Software that has been removed
    foreach ($lastSoftware in $lastChange.Keys) {
        if (-not ($currentChangeLog | Where-Object {$_.Name -eq $lastSoftware})) {
            $softwareChanges += [PSCustomObject]@{
                Name            = $lastSoftware
                FromVersion     = $lastChange[$lastSoftware]
                ToVersion       = 'N/A'
                Status          = "Removed"
            }
        } 
    }

    # Export changes or create a .nochanges file
    if ($softwareChanges.Count -gt 0) {
        $softwareChanges | Export-Csv "$changeLogsFolder\$machine.changelog.$dateTime.csv" -NoTypeInformation
    } else {
        New-Item "$changeLogsFolder\$machine.changelog.$dateTime.nochanges" -ItemType File -Force
    }

    # Remove ALL old *.csvs so that only Change Logs are available
    Get-ChildItem -Path $changeLogsFolder -Filter "*.csv" | Where-Object { $_.Name -like "$machine.startup*" } | Remove-Item
    Write-Host " done!" -NoNewline
}
