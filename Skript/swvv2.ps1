#PSWriteHTML modülünü içe aktarır ~ Skriptin HTML raporları oluşturmak için gerekli modül ~ Dashimo Modulü arsive kaldirildi ve en güncel desteklenen PSWriteHTML

Import-Module PSWriteHTML   #https://github.com/EvotecIT/PSWriteHTML

#---------------------------------------------PATHS--------------------------------------------------------------------------

$BasePath = '\\irt-fss\irt-fs2\SERVICE\IT-HiWi\SWVV2'  #SWVV2 ana dizini

#SWVV skriptinin kullanilacagi bilgisayarlarin listesini computers variable'ina atar
#Bu .txt dosyasindan bilgisayar isimlerini okur ve her satiri bir bilgisayar ismi olarak alir
$ComputerList = Get-Content '$BasePath\Daten\computer_list.txt'

#Windows Versiyon Raporunun kaydedildigi html dosyasinin yolu
$ReportPath = '$BasePath\HTML\WVR.html'

#admin yetkileri icin gerekli xml dosyasinin yolu
$sudo = '$BasePath\Daten\rechte.xml'


#Bilgisayar isimlerini içeren listeyi alır ve çevrimiçi olanları filtreler

$OnlineComputers  = @()
$OfflineComputers = @()

foreach ($Computer in $ComputerList) {
    if (Test-Connection -ComputerName $Computer -Count 1 -Quiet) {
        $OnlineComputers += $Computer
    } else {
        $OfflineComputers += [PSCustomObject]@{
            ComputerName = $Computer
            Status       = 'Offline'
        }
    }
}

#Çevrimiçi bilgisayarlardan Windows sürümü bilgilerini toplar

$WindowsInfo = Invoke-Command -ComputerName $OnlineComputers `
    -Credential $sudo `
    -ThrottleLimit 20 `
    -ErrorAction SilentlyContinue `
    -ScriptBlock {

        $os  = Get-CimInstance Win32_OperatingSystem
        $reg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'

        [PSCustomObject]@{
            ComputerName   = $env:COMPUTERNAME
            OSName         = $os.Caption
            Version        = $os.Version
            Build          = [int]$os.BuildNumber
            DisplayVersion = $reg.DisplayVersion
        }
    }



#HTML raporunu oluşturur ve Windows sürümü bilgilerini tablo olarak ekler
New-HTML -TitleText 'Windows Version Inventory' -FilePath $ReportPath {

    # ONLINE COMPUTERS
    New-HTMLSection -HeaderText 'Online Computers' {

        New-HTMLTable -DataTable $WindowsInfo `
            -Title 'Windows Sürüm Durumu' `
            -Paging `
            -Searching `
            -Sorting `
            -ConditionalFormatting {

                # 🟢 Windows 11 24H2+
                New-HTMLTableCondition -Name 'Build' -Operator ge -Value 26100 -BackgroundColor LightGreen

                # 🟡 Windows 11 23H2
                New-HTMLTableCondition -Name 'Build' -Operator ge -Value 22631 -BackgroundColor Khaki

                # 🔴 Eski / Windows 10
                New-HTMLTableCondition -Name 'Build' -Operator lt -Value 22631 -BackgroundColor Salmon
            }
    }
}