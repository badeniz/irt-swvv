<#Scan Apps & Software of local computers and bring the data to Network as CSV#>

$ErrorActionPreference = "SilentlyContinue"

# Main data path to be given
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

# Add to CSV
$SoftwareList | ConvertTo-Csv -Delimiter ";" -NoTypeInformation | 
    ForEach-Object { $_ -replace '"', '' } | 
    Out-File -FilePath $ExportPath -Encoding UTF8 -Force