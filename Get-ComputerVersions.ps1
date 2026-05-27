<#Scan Apps & Software of local computers and bring the data to Network as CSV#>

$ErrorActionPreference = "SilentlyContinue"

#Enter your own main data path
$DataPath = ""
$ExportPath = "$DataPath\Raw-Data\$($env:COMPUTERNAME).csv"
# 64-bit and 32-bit Registry Paths
$RegPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

 
# Pull apps & software and filter them
$InstalledSoftware = Get-ItemProperty $RegPaths |
    Select-Object DisplayName, DisplayVersion |
    Where-Object { $_.DisplayName -and $_.DisplayName -notmatch "Microsoft|Windows" } |
    Sort-Object -Property DisplayName -Unique

# Determine Windows Version
$WinVer = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion

if (-not $WinVer) { 
    $WinVer = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId 
}

# Add Windows Version to the List
$SoftwareList = [System.Collections.ArrayList]$InstalledSoftware
$SoftwareList.Add([PSCustomObject]@{ DisplayName = "WinVer"; DisplayVersion = $WinVer }) | Out-Null

# Save last execution date
$LastExecuteDate = Get-Date -Format "dd.MM.yyyy HH:mm:ss"

# Add last execution date as if it was a Software itself to the list
$SoftwareList.Add([PSCustomObject]@{ DisplayName = "Scan Date"; DisplayVersion = $LastExecuteDate}) | Out-Null   

# Add to CSV
$SoftwareList | ConvertTo-Csv -Delimiter ";" -NoTypeInformation | 
    ForEach-Object { $_ -replace '"', '' } | 
    Out-File -FilePath $ExportPath -Encoding UTF8 -Force