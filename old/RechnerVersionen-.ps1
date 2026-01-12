<#
Dieses Skript liest alle Registry Einträge von installierter Software auf dem PC aus
und schreibt diese in eine CSV Datei im Temp Ordner
Die Einträge werden vorgefiltert, indem alle Einträge, die "Microsoft" oder "Windows" beinhalten ignoriert werden.
(Pfad: \\irt-fs.win.irt.rwth-aachen.de\IRT\Temp\Bue\RechnerVersionen)
Die Einträge sind nach folgendem Schema aufgebaut:
ProgrammName;Version
Bsp.:
WinVer;1903
Chrome;64.201.24
#>

#Die Einträge sind in der Registry in 64Bit und 32Bit Programme unterteilt.
#hier werden beide ausgelesen und später zusammengefügt
$64Bit = Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Select-Object DisplayName, DisplayVersion | 
    Where-Object {$_ -notmatch ".*Microsoft.*" -and $_ -notmatch ".*Windows.*"}

$32Bit = Get-ItemProperty HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
    Select-Object DisplayName, DisplayVersion |
    Where-Object {$_ -notmatch ".*Microsoft.*" -and $_ -notmatch ".*Windows.*"}

New-PSDrive -Name dest -Root \\irt-fs.win.irt.rwth-aachen.de\IRT\Temp\Bue\RechnerVersionen -PSProvider FileSystem 
    
$64Bit + $32Bit | Sort-Object -Property DisplayName | ConvertTo-Csv -Delimiter ";" -NoTypeInformation | 
    ForEach-Object {$_.Replace('"','')} | Out-File  dest:\$($env:computername).csv
    
#Der Registry Eintrag der Windows Version wird einer für uns Lesbaren Windows Version zugeordnet
function findVersion{
    switch ($args[0]) {
        #Win 10
        10240 { return "1507" }
        10586 { return "1511" }
        14393 { return "1607" }
        15063 { return "1703" }
        16299 { return "1709" }
        17134 { return "1803" }
        17763 { return "1809" }
        18362 { return "1903" }
        18363 { return "1909" }
        19041 { return "2004" }
        #Win 7
        7600 { return "Win7" }
        7601 { return "Win7" }
        Default {return "unknown"}
    }
}

$buildNumber = Get-CimInstance Win32_OperatingSystem | Select-Object BuildNumber
$buildNumber = findVersion $buildNumber."BuildNumber"
"WinVer;$($buildNumber)" >> dest:\$($env:computername).csv