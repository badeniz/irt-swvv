<#Scan Apps & Software of local computers and bring the data to Network as CSV#>

#$ErrorActionPreference = "SilentlyContinue"

# Define paths on the server for the export file
# Add your own Network Path
$DataPath = ""
$ExportPath = "$DataPath\Rawdata\$($env:COMPUTERNAME).csv"

#Registry paths to pull installed software from
$RegPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

# Pull apps & software and filter them to get only the name and version, exclude Microsoft & Windows entries and sort them by name.
$InstalledSoftware = Get-ItemProperty $RegPaths |
    Select-Object DisplayName, DisplayVersion |
    Where-Object { $_.DisplayName -and $_.DisplayName -notmatch "Microsoft|Windows" } |
    Sort-Object -Property DisplayName -Unique

# Determine Windows Version
$WinVer = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion

# If DisplayVersion is not available, fallback to ReleaseId (for older Windows 10 versions)
if (-not $WinVer) { 
    $WinVer = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId 
}

# Disk Space Information
$SystemDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'"
$TotalSizeGB = [math]::Round($SystemDrive.Size / 1GB, 0)
$FreeSpaceGB = [math]::Round($SystemDrive.FreeSpace / 1GB, 0)
$FreePercent = [math]::Round(($SystemDrive.FreeSpace / $SystemDrive.Size) * 100, 0)
$UsedSpaceGB = $TotalSizeGB - $FreeSpaceGB
$UsedPercent = 100 - $FreePercent
$DiskInfoStr = "$FreePercent% Free ($FreeSpaceGB GB free of $TotalSizeGB GB) | $UsedPercent% Used ($UsedSpaceGB GB of $TotalSizeGB GB)"

# Hardware Information (RAM with MHz, CPU, GPU)
$RAM_GB = Get-CimInstance Win32_ComputerSystem | ForEach-Object { [math]::Round($_.TotalPhysicalMemory / 1GB, 0) }
$RAM_MHz = (Get-CimInstance Win32_PhysicalMemory | Select-Object -First 1).ConfiguredClockSpeed
$CPU = (Get-CimInstance Win32_Processor).Name
$GPU = (Get-CimInstance Win32_VideoController | Select-Object -First 1).Name

# Format the RAM string to include MHz info if successfully fetched
$RAM_Str = if ($RAM_MHz) { "${RAM_GB}GB @ ${RAM_MHz}MHz" } else { "${RAM_GB}GB" }
$PCInfoStr = "RAM: $RAM_Str | CPU: $CPU | GPU: $GPU"
# --------------------------------------

#Add Installed Software to an ArrayList
$SoftwareList = [System.Collections.ArrayList]$InstalledSoftware

# Add Windows Version to the List
$SoftwareList.Add([PSCustomObject]@{ DisplayName = "WinVer"; DisplayVersion = $WinVer }) | Out-Null

# Add Disk Info to the List
$SoftwareList.Add([PSCustomObject]@{ DisplayName = "System Drive Info"; DisplayVersion = $DiskInfoStr }) | Out-Null

# Add Hardware Info to the List
$SoftwareList.Add([PSCustomObject]@{ DisplayName = "PC:Info"; DisplayVersion = $PCInfoStr }) | Out-Null

# Save last execution date
$LastExecuteDate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"

# Add last execution date as if it was a Software itself to the list
$SoftwareList.Add([PSCustomObject]@{ DisplayName = "Scan Date"; DisplayVersion = $LastExecuteDate}) | Out-Null   

# Add to CSV
$SoftwareList | ConvertTo-Csv -Delimiter ";" -NoTypeInformation | 
    ForEach-Object { $_ -replace '"', '' } | 
    Out-File -FilePath $ExportPath -Encoding UTF8 -Force