
function Is-ValidVersion($versionString) {
    return $versionString -match '^[0-9]+(\.[0-9]+){1,3}$'
}

<#
Dieses Skript benötigt das Dashimo Modul.
Installationsbefehl:
    Install-Module Dashimo -Force 
https://github.com/EvotecIT/Dashimo
#>

$curPath = Get-Location
$daten_path = "\\irt-fss.irt.rwth-aachen.de\IRT-FS2\SERVICE\SWVV\Daten\"
#$daten_path = "S:\SWVV\Daten"
$versions_path = "$($daten_path)\Compared\Versions.csv"


function checkUpdate{
    <#
    args[0] = Name (z.B. "AdoptOpenJDK")
    args[1] = Link
    args[2] = match String
    args[3] = true/false jenachdem, wie die Seite interpretiert werden soll.
    Wird auf der Seite Code nachgeladen und die HTML Datei somit verändert, wird true übergeben. Ansonsten false, bzw der Parameter weggelassen.
    #>
    $Body = ""
    $counter = 0
    $HTML_TMP = "$($curPath)\" + "$($args[0])TMP.html"
    #Webseite wird aufgerufen und die HTML temporär heruntergeladen.
    if($args[3]){
        $ie = new-object -ComObject "InternetExplorer.Application"
        $url = "$($args[1])"
        $ie.silent = $true
        $ie.navigate($url)
        while($ie.Busy) { Start-Sleep -Milliseconds 1 }
        Start-Sleep 6
        $ie.Document.documentElement.innerHTML > "$($args[0])TMP.html"
        $ie.Stop()
        $ie.Quit()
    }   
    else {
        $wc=new-object system.net.webclient
        $wc.UseDefaultCredentials = $true
        $wc.downloadfile($args[1] , $HTML_TMP)
    }
    #HTML wird nach match String durchsucht und die aktuelle Version ausgelesen.
    #String wird in CSV gültigem Format abgespeichert.
    #Bsp.: "AdoptOpenJDK;4.1.0"
    foreach($line in Get-Content $HTML_TMP){
        if($line -match $args[2]){
            $body = "$($args[0]);$($matches['build'])"
            $counter = 1
            break
        }
    }
    #Temporäre HTML Datei wird gelöscht
    Remove-Item -path $HTML_TMP
    #Sollte der match String nicht gefunden werden wird eine Error Meldung eingetragen
    if($counter -eq 0){
        "$($args[0]);NotFound"
    }

    return $Body
}

#Aufruf der Verschiedenen Webseiten
#Build Strings werden in $output abgespeichert

$output = @("DisplayName;DisplayVersion")

#Eclipse Temurin JDK
$output += "Eclipse Temurin JDK mit Hotspot;8.0.302.8"
#$output += checkUpdate "AdoptOpenJDK" "https://adoptopenjdk.net/release_notes.html?variant=openjdk11&jvmVariant=hotspot" ".*<a id=""jdk11_.*"">OpenJDK (?<build>.*) \(.*</a></h2>"

#AdobeReader
$output += "Adobe Acrobat Reader;21.007.20099"

#GhostScript
$output += "GPL GhostScript;9.55.0"
#$output += checkUpdate "GPL GhostScript" "https://www.ghostscript.com/releases.html" "(The latest AGPL release is <a href=""Ghostscript).*(.html"">Ghostscript )(?<build>.*)(</a>.*.</p>)"

#JabRef 
$output += checkUpdate "JabRef" "https://www.fosshub.com/JabRef.html" "<a class=""fSignature"" href=""\#"" data-file=""JabRef-(?<build>.*).msi"" rel=""nofollow"">Signature</a>"

#MiKTeX
$output += checkUpdate "MiKTeX" "https://miktex.org/download" "        <div class=.*>basic-miktex-(?<build>.*)-x64.exe</div>"

#TeXnicCenter
$output += checkUpdate "TeXnicCenter" "http://www.texniccenter.org/download/" "(	<li><a href=""http://sourceforge.net/projects/texniccenter/files/TeXnicCenter/)(.*)(Stable/TXCSetup_)(.*)(Stable_Win32.exe/download"">TeXnicCenter )(?<build>.*)( Stable .*32(.*)</a></li>)" $true

#TortoiseGit
$output += checkUpdate "TortoiseGit" "https://tortoisegit.org/download/" ".*download.tortoisegit.org/tgit/(?<build>.*)/</a>.*"

#NotePad++
$output += checkUpdate "NotePad" "https://notepad-plus-plus.org/download/" "(          <a href=""/downloads/.*/""><strong>   )(?<build>.*)(</strong></a>)" $true

#7-Zip
$output += checkUpdate "7-Zip" "https://www.7-zip.org/index.html" "<p><b>Download 7-Zip (?<build>.*) \(20.*\) for Windows</b>:</p> "

#VLC-Player
$output += checkUpdate "VLC media player" "https://www.videolan.org/vlc/index.de.html" "            <a id='downloadButton2' class=""btn btn-default btn-lg btn-dl"" href='//get.videolan.org/vlc/(?<build>.*)/win32.*>"

#WinVer
$output += "WinVer;24H2"

#Firefox
$FFLatest = ( Invoke-WebRequest  "https://product-details.mozilla.org/1.0/firefox_versions.json" -UseBasicParsing | ConvertFrom-Json ).LATEST_FIREFOX_VERSION
$output += "Firefox;$($FFLatest)"
#$output += checkUpdate "Firefox" "https://caniuse.com/" '<span class="home__browser-bar  home__browser-bar--full  browser-bg--firefox" title="Full support: .*pts" style="width: 82.2%;">Firefox(?<build>.*): .*</span><span'
#$output += checkUpdate "Firefox;113" 

#Chrome / Code Update 07.06.2023
#    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#   $j = Invoke-WebRequest 'https://omahaproxy.appspot.com/all.json' | ConvertFrom-Json
#    $output += "Google Chrome;$($j[0].versions[4].current_version)"
$output += checkUpdate "Google Chrome" "https://chromeenterprise.google/intl/de_DE/browser/download/#windows-tab" "<span class=""ce-download__download-version-number"">(?<build>.*)</span>" $true
#$output += checkUpdate "Google Chrome;114.0.5735"



#Citavi
$output += "Citavi;	6.10.0.0"
#$output += checkUpdate "Citavi" "https://help.citavi.com/en/knowledge-base/article/citavi-version-history" "<li> Citavi (?<build>.*) (update) , 2021-01-14, <a href="https://www.citavi.com/en/news/articles/update-to-citavi-6-8-now-available">improvements</a>, <a href="https://setup1.citavi.com/release/default/Citavi6Setup.exe">download</a></li>  "

#Ghostview
$output += "GSview;5.0"

#InkScape
$output += "Inkscape;1.1.1"

#CorelDraw
$output += "CorelDRAW;25.0"

#Python
$output += "Python;	3.10.7581.0"


#Sciebo
$output += "Sciebo;	2.8.2.4365"

#$update += checkUpdate "Sciebo" "https://hochschulcloud.nrw/de/download/index.html" "<a class="btn cs-btn black download-btn" href="https://www.sciebo.de/install/sciebo-2.7.6.3308.x64.msi" role="button" lang="en">Download<img src="/assets/images/ui/arrow.svg" class="img-fluid" alt=""></a>"

#VisualStudioCommunity
$output += "Visual Studio Community 2019;16.11.7"


#Matlab
$output += "MATLAB 2024a;24.1"

#Office
$output += "Microsoft Office Professional Plus;16.0.10382.20010"


#Finde Versions.csv
$path = Get-Content $versions_path

#Sollte die Datei bereits existieren wird sie gelöscht.
if($null -ne $path){
    Remove-Item -path $versions_path
}

#Danach wird der inhalt von $output in Versions.csv geschrieben
#Wir erhalten somit eine saubere CSV Datei, die die Namen und aktuellen Versionsstände der zu überwachenden Software beinhaltet.
foreach($line in $output){
    #args[1] = Output Path
    $line >> $versions_path
}


#Compare Versions 
$versions = Import-Csv $versions_path -Delimiter ";"
$control = @()

#Liste von Toleranzen für Versionsnummern
$jdk = @(0,1,0,0)
$adobe = @(0,1,0,0)
$gpl = @(0,1,0,0)
$firefox = @(0,1,0,0)
$miktex = @(1,1,1,1)
$tex = @(0,1,0,0)
$chrome = @(0,0,0,1)
$git = @(0,1,0,0)
$notepad = @(0,1,0,0)
$7zip = @(0,1,1,1)
$vlc = @(0,1,0,0)
$citavi = @(0,1,0,0)
$inkscape = @(0,1,0,0)
$ghostscript = @(0,1,0,0)
$coreldraw = @(0,1,0,0)
$python = @(0,1,0,0)
$sciebo = @(0,1,0,0)
$visualstudio = @(0,1,0,0)
$win = @("24H2","23H2")
$matlab = @(0,1,1,0)
$office = @(0,1,0,0)


#Für jede PC.csv
#Vergleiche die Versionen auf den Computern, mit den aktuellen aus Versions.csv
Get-ChildItem "$($daten_path)\Rohdaten" -Filter *.csv |
ForEach-Object {
    $out_of_date = $false
    
    $output = @("DisplayName;DisplayVersion;NewVersion;Update")
    $computer = Import-Csv $_.FullName -Delimiter ";"
    $list = @( @("Adobe Acrobat Reader", $adobe), @("Eclipse Temurin JDK mit Hotspot", $jdk), @("Citavi", $citavi),  
            @("CorelDraw", $coreldraw),@("GSview", $ghostview), @("Google Chrome", $chrome), @("GPL GhostScript", $gpl),  @("InkScape", $inkscape), 
            @("Matlab", $matlab), @("Microsoft Office", $office), @("MiKTeX", $miktex), @("Firefox", $firefox), @("NotePad", $notepad),  
            @("Python", $python), @("Sciebo", $sciebo), @("TeXnicCenter", $tex), @("TortoiseGit", $git), @("Visual Studio Community", $visualstudio), 
            @("VLC media player", $vlc), @("7-Zip", $7zip), @("WinVer", $win))
              


    #Für jedes zu überwachende Programm werden die Versionsstände des Computers
    #und die aktuellste Versionsstände aus der jeweiligen CSV ausgelesen.
    foreach($object in $list){
        $version_Computer = $computer | Where-Object {$_ -match ".*$($object[0]).*"}
        $version_Online = $versions | Where-Object {$_ -match ".*$($object[0]).*"}

        if($object -eq "WinVer"){
            if($computer."DisplayVersion" -ge $object[1][0]){
                $output += "WinVer;$($version_computer."DisplayVersion");;UTD"
            } else {
                if($computer."DisplayVersion" -ge $object[1][1]){
                    $output += "WinVer;$($version_computer."DisplayVersion");;SOOD"
                } else {
                    $output += "WinVer;$($version_computer."DisplayVersion");;OOD"
                }
            }
            $lastWrite = $_.LastWriteTime
            $timespan = new-timespan -days 30
            if (((get-date) - $lastWrite) -gt $timespan) {
                $output += "Daten von;$lastWrite;;OOD"
            } else {
                $output += "Daten von;$lastWrite;;UTD"
            }            
        }
        #Sollte es mehrere passende Zeilen geben
        #scheint die Zweite Zeile soweit immer passend 
        #zu sein.
        if($version_Computer -is [Array]){  
            if($version_Computer."DisplayVersion"[-1] -eq $version_Online."DisplayVersion" -or [version]$version_Computer."DisplayVersion"[-1] -ge [version]$version_Online."DisplayVersion"){
                $output += "$($version_Computer."DisplayName"[-1]);$($version_Computer."DisplayVersion"[-1]);$($version_Online."DisplayVersion");UTD"
            } else {
                #Minor Version wird um 1 erhöht.
                #Ist die Buildnummer nicht mehr kleiner ist das Programm nur leicht out of date
                $v = [version]$version_Computer."DisplayVersion"[1]
                if($v.Revision -ne -1){
                    $vUp = [version]::New($v.Major+$object[1][0],$v.Minor+$object[1][1],$v.Build+$object[1][2],$v.Revision+$object[1][3])
                } else {
                    $vUp = [version]::New($v.Major+$object[1][0],$v.Minor+$object[1][1],$v.Build+$object[1][2])
                }
                if($vUp -eq $version_Online."DisplayVersion" -or $vUp -ge [version]$version_Online."DisplayVersion"){
                    $output += "$($version_Computer."DisplayName"[-1]);$($version_Computer."DisplayVersion"[-1]);$($version_Online."DisplayVersion");SOOD"
                } else {
                    $output += "$($version_Computer."DisplayName"[-1]);$($version_Computer."DisplayVersion"[-1]);$($version_Online."DisplayVersion");OOD"
                    $out_of_date = $true
                }
            }
        } else {
            if($version_Computer -ne $null){
                if($version_Computer."DisplayVersion".Split(" ")[0] -eq $version_Online."DisplayVersion"){
                    $output += "$($version_Computer."DisplayName");$($version_Computer."DisplayVersion");$($version_Online."DisplayVersion");UTD"
                }
                else{
                     #Minor Version wird um 1 erhöht.
                    #Ist die Buildnummer nicht mehr kleiner ist das Programm nur leicht out of date
                    $v = [version]$version_Computer."DisplayVersion".Split(" ")[0]
                    if($v.Revision -ne -1){
                        $vUp = [version]::New($v.Major+$object[1][0],$v.Minor+$object[1][1],$v.Build+$object[1][2],$v.Revision+$object[1][3])
                    } else {
                        $vUp = [version]::New($v.Major+$object[1][0],$v.Minor+$object[1][1],$v.Build+$object[1][2])
                    }
                    if($vUp -eq $version_Online."DisplayVersion" -or $vUp -ge [version]$version_Online."DisplayVersion"){
                        $output += "$($version_Computer."DisplayName");$($version_Computer."DisplayVersion");$($version_Online."DisplayVersion");SOOD"
                    } else {
                        $output += "$($version_Computer."DisplayName");$($version_Computer."DisplayVersion");$($version_Online."DisplayVersion");OOD"
                        $out_of_date = $true
                    }
                }
            }
            else{
                #Sollte $version_Computer == $null sein, ist das Programm nicht auf dem Computer installiert.
                $output += "$($object[0]);Nicht installiert;$($version_Online."DisplayVersion");"
            }
        }
    }
    
    if($out_of_date){
        #$control sammelt alle Computernamen, die veraltete Software enthalten
        $control += $_.BaseName
    }
    $path = Get-Content "$($daten_path)Compared\CSV\$_"
    if($null -ne $path){
        Remove-Item -path "$($daten_path)Compared\CSV\$_"
    }
    #Danach werden die verglichenen Programme in einer neuen CSV datei temporär abgespeichert.
    foreach($line in $output){
        $line >> "$($daten_path)Compared\CSV\$_"
    }
}

#Hier wird der Inhalt von Control übertragen.
$path = Get-Content "$($daten_path)\Rohdaten\needUpdate.txt"
if($null -ne $path){
   Remove-Item -path "$($daten_path)\Rohdaten\needUpdate.txt"
}
foreach($line in $control){
    $line >> "$($daten_path)\Rohdaten\needUpdate.txt"
}

#Write HTML Files
#Für jeden Computer wird eine HTML datei angelegt, in der die Programme und Versionsstände in einer Tabelle anschaulich dargestelt werden.
Get-ChildItem "$($daten_path)\Compared\CSV" -Filter *.csv |
ForEach-Object {
    $Process = Import-Csv $_.FullName -Delimiter ";"

    Dashboard -Name $_.BaseName -FilePath  "$($daten_path)\HTML\$($_.BaseName).html"{
        Tab -Name $_.BaseName {
            Section -Name $_.BaseName {
                Table -DataTable $Process {
                    TableConditionalFormatting -Name 'Update' -ComparisonType string -Operator eq -Value 'UTD' -Color White -BackgroundColor Green -Row
                    TableConditionalFormatting -Name 'Update' -ComparisonType string -Operator eq -Value 'OOD' -Color White -BackgroundColor Crimson -Row
                    TableConditionalFormatting -Name 'Update' -ComparisonType string -Operator eq -Value 'SOOD' -Color White -BackgroundColor Orange -Row
                }
            }
        }
    }
}

#Delete \Compared\CSV
#Alle temporären Dateien werden gelöscht
<#Get-ChildItem "$($daten_path)\Compared\CSV" -Filter *.csv |
ForEach-Object {
    Remove-Item -Path \\irt-fs2.irt.rwth-aachen.de\irt2\service\SWVV\Daten\Compared\CSV\$_
}
#>

#HTML Einträge, für eine Übersichtsdatei
$needsUpdate = Get-Content "$($daten_path)\Rohdaten\needUpdate.txt"
$content = @()
$content += "<!DOCTYPE html>"
$content += '<html lang="en">'
$content += "<head>"
$content += '<meta charset="UTF-8">'
$content += '<meta http-equiv="X-UA-Compatible" content="IE=edge">'
$content += '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
$content += '<link rel="stylesheet" href="styles.css">'
$content += "<title>Rechneruebersicht</title>"
$content += "</head>"
$content += "<body>"
$content += "<h1>Rechner Updates</h1>"
$content += "<div class=""links-block"">"



#create Overview.html
#In Overview.html werden alle HTML Dateien zu den einzelnen Computern verlinkt.
#Computer, die veraltete Software haben werden rot hinterlegt
#aktuelle Computer werden grün hinterlegt
Get-ChildItem "$($daten_path)\HTML" -Filter *.html |
ForEach-Object {
    $unused = $true
    foreach($line in $needsUpdate){
        if($line -eq $_.BaseName){
            $content += "<a href=""$($_.Name)"" target=""_blank"" class=""red"" id=""$($_.BaseName)"">$($_.BaseName)</a><br>"
            $unused = $false
        }
    }
    if($unused){
        $content += "<a href=""$($_.Name)"" target=""_blank"" class=""green"" id=""$($_.BaseName)"">$($_.BaseName)</a><br>"
    }
}
#Sollte eine alte Overview.html existieren wird diese Gelöscht 
$path = Get-Content "$($daten_path)\HTML\Overview.html"
if($null -ne $path){
    Remove-Item -path "$($daten_path)\HTML\Overview.html"
}
#und dann neu angelegt.
foreach($line in $content){
    $line >> "$($daten_path)\HTML\Overview.html"
}

#html abschliessen
$content += "</div>"
$content += "</body>">> "$($daten_path)\HTML\Overview.html"
$content += "</html>">> "$($daten_path)\HTML\Overview.html"

 


#CSS erstellen
#$daten_path = "\\irt-fs2.irt.rwth-aachen.de\irt2\service\SWVV\Daten\"
#$path = Get-Content "$($daten_path)\HTML\styles.css"
#if($null -ne $path){
#    Remove-Item -path "$($daten_path)\HTML\styles.css"
#}
#"h1 {color: red}">> "$($daten_path)\HTML\styles.css"

