<#Compares app & software versions from web or our static list and produces native HTML reports#>

# Main data path to be given
$DataPath = ""
$RawDataPath = "$DataPath\Raw-Data"
$HTMLPath = "$DataPath\HTML"

if (-not (Test-Path $HTMLPath)) { New-Item -ItemType Directory -Path $HTMLPath -Force | Out-Null }
if (-not (Test-Path $RawDataPath)) { New-Item -ItemType Directory -Path $RawDataPath -Force | Out-Null }

# Function to pull versions from web
function Get-WebVersion {
    param([string]$Url, [string]$RegexMatch, [string]$RegexGroup = "build")
    try {
        $Web = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 10
        if ($Web.Content -match $RegexMatch) { return $Matches[$RegexGroup] }
    } catch { return "Not found" }
}

# Determine updated versions
$Latest = @{}

# Firefox API
$Latest["Firefox"] = (Invoke-RestMethod "https://product-details.mozilla.org/1.0/firefox_versions.json").LATEST_FIREFOX_VERSION

# Chrome API
$ChromeAPI = Invoke-RestMethod "https://versionhistory.googleapis.com/v1/chrome/platforms/win/channels/stable/versions"
$Latest["Google Chrome"] = $ChromeAPI.versions[0].version

# Those to pull from web with regex

#7-Zip
$Latest["7-Zip"] = Get-WebVersion -Url "https://www.7-zip.org/index.html" -RegexMatch "Download 7-Zip (?<build>[0-9\.]+) \(20"

#VLC media player
$Latest["VLC media player"] = Get-WebVersion -Url "https://www.videolan.org/vlc/index.de.html" -RegexMatch "get\.videolan\.org/vlc/(?<build>[0-9\.]+)/win32"

#MiKTeX
$Latest["MiKTeX"] = Get-WebVersion -Url "https://miktex.org/download" -RegexMatch "        <div class=.*>basic-miktex-(?<build>.*)-x64.exe</div>"

#TeXnicCenter
$Latest["TeXnicCenter"] = Get-WebVersion -Url "http://www.texniccenter.org/download/" -RegexMatch "(	<li><a href=""http://sourceforge.net/projects/texniccenter/files/TeXnicCenter/)(.*)(Stable/TXCSetup_)(.*)(Stable_Win32.exe/download"">TeXnicCenter )(?<build>.*)( Stable .*32(.*)</a></li>)" -RegexGroup "build"

#TortoiseGit
$Latest["TortoiseGit"] = Get-WebVersion -Url "https://tortoisegit.org/download/" -RegexMatch ".*download.tortoisegit.org/tgit/(?<build>.*)/</a>.*"

#NotePad++
$Latest["NotePad++"] = Get-WebVersion -Url "https://notepad-plus-plus.org/downloads/" -RegexMatch 'Current Version (?<build>[0-9\.]+)'

# Static definitions to save time and update manually
$Latest["Adobe Acrobat Reader"] = "21.007.20099"
$Latest["GPL GhostScript"]      = "9.55.0"
$Latest["Eclipse Temurin JDK"]  = "8.0.302.8"
$Latest["Citavi"]               = "6.10.0.0"
$Latest["GSview"]               = "5.0"
$Latest["Inkscape"]             = "1.1.1"
$Latest["CorelDRAW"]            = "25.0"
$Latest["Python"]               = "3.10.7581.0"
$Latest["Sciebo"]               = "2.8.2.4365"
$Latest["Visual Studio Community 2019"] = "16.11.7"
$Latest["MATLAB 2024a"]         = "24.1"
$Latest["Microsoft Office Professional Plus"] = "16.0.10382.20010"
$Latest["WinVer"]               = "25H2"
# Add as you wish

#HTML CSS and Style Definitions
$HTML_Head = @"
<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Computer Reports</title>
<style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f4f4f9;}
    table { border-collapse: collapse; width: 100%; box-shadow: 0 2px 5px rgba(0,0,0,0.1); background: #fff; }
    th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
    th { background-color: #2c3e50; color: white; }
    .UTD { background-color: #c3e6cb; } /* Updated - Green */
    .SOOD { background-color: #ffeeba; } /* A bit outdated - Orange */
    .OOD { background-color: #f5c6cb; } /* Outdated - Red */
    .links a { display: block; padding: 5px; text-decoration: none; font-weight: bold; }
    .links a.OOD { color: #dc3545; }
    .links a.UTD { color: #28a745; }
</style></head><body>
"@

$OverviewLinks = @()

# Comparing the Computers
Get-ChildItem -Path $RawDataPath -Filter "*.csv" | ForEach-Object {
    $PCName = $_.BaseName
    $PCData = Import-Csv $_.FullName -Delimiter ";"
    $ReportData = @()
    $HasOutdated = $false

    $DateRow = $PCData | Where-Object { $_.DisplayName -match "Scan Date" } | Select-Object -Last 1
    $LastScan = if ($DateRow.DisplayVersion) { $DateRow.DisplayVersion.Trim() } else { "Unknown / Old Date" }
    

    foreach ($SoftName in $Latest.Keys) {
        $Installed = $PCData | Where-Object { $_.DisplayName -match [regex]::Escape($SoftName) } | Select-Object -Last 1
        $LatestVer = $Latest[$SoftName]
        $StatusInfo = "Not Found"
        
        if ($Installed) {
            $InstVerStr = $Installed.DisplayVersion -replace '[^\d\.]', '' # Sadece rakam ve nokta
            
            try {
                $vInst = [version]$InstVerStr
                $vLate = [version]($LatestVer -replace '[^\d\.]', '')
                
                if ($vInst -ge $vLate) { 
                    $StatusInfo = "UTD" 
                } elseif ($vInst.Major -lt $vLate.Major) { 
                    $StatusInfo = "OOD"; $HasOutdated = $true 
                } else { 
                    $StatusInfo = "SOOD"; $HasOutdated = $true 
                }
            } catch {
                if ($Installed.DisplayVersion -eq $LatestVer) { $StatusInfo = "UTD" } else { $StatusInfo = "OOD"; $HasOutdated = $true }
            }

            $ReportData += [PSCustomObject]@{
                Software = $SoftName
                LoadedVersion = $Installed.DisplayVersion
                UpdatedVersion = $LatestVer
                Status = $StatusInfo
            }
        } else {
            $ReportData += [PSCustomObject]@{ Software = $SoftName; LoadedVersion = "Not downloaded"; UpdatedVersion = $LatestVer; Status = "Not Found" }
        }
    }

    # Convert Table to HTML
    $TableHTML = $ReportData | ConvertTo-Html -Fragment | Out-String
    $TableHTML = $TableHTML -replace '<tr><td>(.*?)</td><td>(.*?)</td><td>(.*?)</td><td>UTD</td></tr>', '<tr class="UTD"><td>$1</td><td>$2</td><td>$3</td><td>Updated</td></tr>'
    $TableHTML = $TableHTML -replace '<tr><td>(.*?)</td><td>(.*?)</td><td>(.*?)</td><td>SOOD</td></tr>', '<tr class="SOOD"><td>$1</td><td>$2</td><td>$3</td><td>Minor update available</td></tr>'
    $TableHTML = $TableHTML -replace '<tr><td>(.*?)</td><td>(.*?)</td><td>(.*?)</td><td>OOD</td></tr>', '<tr class="OOD"><td>$1</td><td>$2</td><td>$3</td><td>Old</td></tr>'
    $TableHTML = $TableHTML -replace '<tr><td>(.*?)</td><td>(.*?)</td><td>(.*?)</td><td>Not Found</td></tr>', '<tr><td>$1</td><td>$2</td><td>$3</td><td>Not downloaded</td></tr>'


    # HTML Save according to PC
    $FinalHTML = $HTML_Head + "<h2>$PCName Update Status</h2>"
    $FinalHTML += "<p style='color: #0099ff; font-size: 14px;'><strong>Last Scan:</strong> $LastScan</p>"
    $FinalHTML += $TableHTML + "</body></html>"
    
    $FinalHTML | Set-Content -Path "$HTMLPath\$PCName.html" -Encoding UTF8

    # Saving (Overview) Links
    $OverviewLinks += [PSCustomObject]@{ Name = $PCName; IsOutdated = $HasOutdated }
}

# Creating Overview file as HTML
$OverviewHTML = $HTML_Head + "<h1>All Computers</h1><div class='links'>"
foreach ($Link in ($OverviewLinks | Sort-Object Name)) {
    $Class = if ($Link.IsOutdated) { "OOD" } else { "UTD" }
    $OverviewHTML += "<a class='$Class' href='$($Link.Name).html' target='_blank'>$($Link.Name)</a>"
}
$OverviewHTML += "</div></body></html>"
$OverviewHTML | Set-Content -Path "$HTMLPath\Overview.html" -Encoding UTF8

Write-Host "Process completed, script invoked. Reports are to be found at $HTMLPath ." -ForegroundColor Green