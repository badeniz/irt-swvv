<#
Dieses Skript benötigt das Dashimo Modul.
Installationsbefehl:
    Install-Module Dashimo -Force 
https://github.com/EvotecIT/Dashimo
#>

$curPath = Get-Location

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

#AdoptOpenJDK
$output += checkUpdate "AdoptOpenJDK" "https://adoptopenjdk.net/release_notes.html?variant=openjdk11&jvmVariant=hotspot" ".*<a id=""jdk11_.*"">OpenJDK (?<build>.*) \(.*</a></h2>"

#AdobeReader
$output += checkUpdate "Adobe Acrobat Reader" "https://get.adobe.com/de/reader/" "(			mainInstallerName: ""Reader DC )20(?<build>.*)( German for Windows"",)" $true

#GhostScript
$output += checkUpdate "GPL GhostScript" "https://www.ghostscript.com/releases.html" "(The latest AGPL release is <a href=""Ghostscript).*(.html"">Ghostscript )(?<build>.*)(</a>.*.</p>)"

#JabRef 
$output += checkUpdate "JabRef" "https://www.fosshub.com/JabRef.html" "<a class=""fSignature"" href=""\#"" data-file=""JabRef-(?<build>.*).msi"" rel=""nofollow"">Signature</a>"

#MiKTeX
$output += checkUpdate "MiKTeX" "https://miktex.org/download" "        <div class=.*>basic-miktex-(?<build>.*)-x64.exe</div>"

#TeXnicCenter
$output += checkUpdate "TeXnicCenter" "http://www.texniccenter.org/download/" "(	<li><a href=""http://sourceforge.net/projects/texniccenter/files/TeXnicCenter/)(.*)(Stable/TXCSetup_)(.*)(Stable_Win32.exe/download"">TeXnicCenter )(?<build>.*)( Stable .*32(.*)</a></li>)" $true

#TortoiseGit
$output += checkUpdate "TortoiseGit" "https://tortoisegit.org/download/" ".*download.tortoisegit.org/tgit/(?<build>.*)/</a>.*"

#NotePad++
$output += checkUpdate "NotePad" "https://notepad-plus-plus.org/download/" "(          <a href=""/downloads/.*/""><strong>Current Version )(?<build>.*)(</strong></a>)" $true

#7-Zip
$output += checkUpdate "7-Zip" "https://www.7-zip.org/index.html" "<p><b>Download 7-Zip (?<build>.*) \(20.*\) for Windows</b>:</p> "

#VLC-Player
$output += checkUpdate "VLC media player" "https://www.videolan.org/vlc/index.de.html" "            <a id='downloadButton2' class=""btn btn-default btn-lg btn-dl"" href='//get.videolan.org/vlc/(?<build>.*)/win32.*>"

#WinVer
$output += "WinVer;1903"

#Firefox
$FFLatest = ( Invoke-WebRequest  "https://product-details.mozilla.org/1.0/firefox_versions.json" -UseBasicParsing | ConvertFrom-Json ).LATEST_FIREFOX_VERSION
$output += "Firefox;$($FFLatest)"

#Chrome
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $j = Invoke-WebRequest 'https://omahaproxy.appspot.com/all.json' | ConvertFrom-Json
    $output += "Google Chrome;$($j[0].versions[4].current_version)"

#Finde Versions.csv
$path = Get-Content T:\Bue\RechnerVersionen\Compared\Versions.csv

#Sollte die Datei bereits existieren wird sie gelöscht.
if($null -ne $path){
    Remove-Item -path T:\Bue\RechnerVersionen\Compared\Versions.csv
}

#Danach wird der inhalt von $output in Versions.csv geschrieben
#Wir erhalten somit eine saubere CSV Datei, die die Namen und aktuellen Versionsstände der zu überwachenden Software beinhaltet.
foreach($line in $output){
    #args[1] = Output Path
    $line >> T:\Bue\RechnerVersionen\Compared\Versions.csv
}

#Compare Versions 
$versions = Import-Csv T:\Bue\RechnerVersionen\Compared\Versions.csv -Delimiter ";"
$control = @()

#Für jede PC.csv
#Vergleiche die Versionen auf den Computern, mit den aktuellen aus Versions.csv
Get-ChildItem "T:\Bue\RechnerVersionen" -Filter *.csv |
ForEach-Object {
    $out_of_date = $false
    
    $output = @("DisplayName;DisplayVersion;NewVersion;Update")
    $computer = Import-Csv $_.FullName -Delimiter ";"
    $list = @("AdoptOpenJDK", "Adobe Acrobat Reader", "GPL GhostScript",
              "Firefox", "MiKTeX", "TeXnicCenter", "Google Chrome",
              "TortoiseGit", "NotePad", "7-Zip",
              "VLC media player", "WinVer")

    #Für jedes zu überwachende Programm werden die Versionsstände des Computers
    #und die aktuellste Versionsstände aus der jeweiligen CSV ausgelesen.
    foreach($object in $list){
        $version_Computer = $computer | Where-Object {$_ -match ".*$($object).*"}
        $version_Online = $versions | Where-Object {$_ -match ".*$($object).*"}

        #Sollte es mehrere passende Zeilen geben
        #scheint die Zweite Zeile soweit immer passend 
        #zu sein.
        if($version_Computer -is [Array]){
            if($version_Computer."DisplayVersion"[1] -eq $version_Online."DisplayVersion" -or [version]$version_Computer."DisplayVersion"[1] -gt [version]$version_Online."DisplayVersion"){
                $output += "$($version_Computer."DisplayName"[1]);$($version_Computer."DisplayVersion"[1]);$($version_Online."DisplayVersion");UTD"
            }
            else{
                $output += "$($version_Computer."DisplayName"[1]);$($version_Computer."DisplayVersion"[1]);$($version_Online."DisplayVersion");OOD"
                $out_of_date = $true
            }
        }     
        else {
            if($version_Computer -ne $null){
                if($version_Computer."DisplayVersion".Split(" ")[0] -eq $version_Online."DisplayVersion" -or [version]$version_Computer."DisplayVersion".Split(" ")[0] -gt [version]$version_Online."DisplayVersion"){
                    $output += "$($version_Computer."DisplayName");$($version_Computer."DisplayVersion");$($version_Online."DisplayVersion");UTD"
                }
                else{
                    $output += "$($version_Computer."DisplayName");$($version_Computer."DisplayVersion");$($version_Online."DisplayVersion");OOD"
                    $out_of_date = $true
                }
            }
            else{
                #Sollte $version_Computer == $null sein, ist das Programm nicht auf dem Computer installiert.
                $output += "$($object);Nicht installiert;$($version_Online."DisplayVersion");"
            }
        }
    }
    if($out_of_date){
        #$control sammelt alle Computernamen, die veraltete Software enthalten
        $control += $_.BaseName
    }
    #Danach werden die verglichenen Programme in einer neuen CSV datei temporär abgespeichert.
    foreach($line in $output){
        $line >> T:\Bue\RechnerVersionen\Compared\CSV\$_
    }
}

#Hier wird der Inhalt von Control übertragen.
$path = Get-Content T:\Bue\RechnerVersionen\needUpdate.txt
if($null -ne $path){
   Remove-Item -path T:\Bue\RechnerVersionen\needUpdate.txt
}
foreach($line in $control){
    $line >> T:\Bue\RechnerVersionen\needUpdate.txt
}

#Write HTML Files
#Für jeden Computer wird eine HTML datei angelegt, in der die Programme und Versionsstände in einer Tabelle anschaulich dargestelt werden.
Get-ChildItem "T:\Bue\RechnerVersionen\Compared\CSV" -Filter *.csv |
ForEach-Object {
    $Process = Import-Csv $_.FullName -Delimiter ";"

    Dashboard -Name $_.BaseName -FilePath $PSScriptRoot\$($_.BaseName).html{
        Tab -Name $_.BaseName {
            Section -Name $_.BaseName {
                Table -DataTable $Process {
                    TableConditionalFormatting -Name 'Update' -ComparisonType string -Operator eq -Value 'UTD' -Color White -BackgroundColor Green -Row
                    TableConditionalFormatting -Name 'Update' -ComparisonType string -Operator eq -Value 'OOD' -Color White -BackgroundColor Crimson -Row
                }
            }
        }
    }
}

#Delete \Compared\CSV
#Alle temporären Dateien werden gelöscht
Get-ChildItem "T:\Bue\RechnerVersionen\Compared\CSV" -Filter *.csv |
ForEach-Object {
    Remove-Item -Path T:\Bue\RechnerVersionen\Compared\CSV\$_
}

#HTML Einträge, für eine Übersichtsdatei
$needsUpdate = Get-Content T:\Bue\RechnerVersionen\needUpdate.txt
$content = @()
$content += "<html>"
$content += "<title>Rechneruebersicht</title>"
$content += "</head>"
$content += "<body>"
$content += "<h1>Rechner Updates</h1>"


#create Overview.html
#In Overview.html werden alle HTML Dateien zu den einzelnen Computern verlinkt.
#Computer, die veraltete Software haben werden rot hinterlegt
#aktuelle Computer werden grün hinterlegt
Get-ChildItem "T:\Bue\RechnerVersionen\Compared\CSV" -Filter *.html |
ForEach-Object {
    $unused = $true
    foreach($line in $needsUpdate){
        if($line -eq $_.BaseName){
            $content += "<a href=""$($_.FullName)""><font color=""FF0000"">$($_.BaseName)</font></a><br>"
            $unused = $false
        }
    }
    if($unused){
        $content += "<a href=""$($_.FullName)""><font color=""00FF00"">$($_.BaseName)</font></a><br>"
    }
}
#Sollte eine alte Overview.html existieren wird diese Gelöscht 
$path = Get-Content T:\Bue\RechnerVersionen\Compared\CSV\Overview.html
if($null -ne $path){
    Remove-Item -path T:\Bue\RechnerVersionen\Compared\CSV\Overview.html
}
#und dann neu angelegt.
foreach($line in $content){
    $line >> T:\Bue\RechnerVersionen\Compared\CSV\Overview.html
}