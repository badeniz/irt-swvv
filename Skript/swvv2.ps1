#PSWriteHTML modülünü içe aktarır ~ Skriptin HTML raporları oluşturmak için gerekli modül ~ Dashimo Modulü arsive kaldirildi ve en güncel desteklenen PSWriteHTML
Import-Module PSWriteHTML   #https://github.com/EvotecIT/PSWriteHTML

Write-Host "PSWriteHTML modulu yuklendi." -ForegroundColor Green

#---------------------------------------------PATHS--------------------------------------------------------------------------

$BasePath = '\\irt-fss\irt-fs2\SERVICE\IT-HiWi\SWVV2'  #SWVV2 ana dizini
$ComputerList = Get-Content "$BasePath\Daten\computer_list.txt"
$ReportPath = "$BasePath\HTML\WVR.html"

#---------------------------------------------INIT-------------------------------------------------------------------------

$Results = @()
$ExportTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

#---------------------------------------------COLLECT DATA-----------------------------------------------------------------

foreach ($Computer in $ComputerList) {

    Write-Host "Kontrol ediliyor: $Computer" -NoNewline

    try {
        if ($Computer -ieq $env:COMPUTERNAME) {
            $os  = Get-CimInstance Win32_OperatingSystem
            $reg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
        }
        else {
            $data = Invoke-Command -ComputerName $Computer -ErrorAction Stop -ScriptBlock {
                $os  = Get-CimInstance Win32_OperatingSystem
                $reg = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'

                [PSCustomObject]@{
                    Caption        = $os.Caption
                    Version        = $os.Version
                    Build          = [int]$os.BuildNumber
                    DisplayVersion = $reg.DisplayVersion
                }
            }
            $os  = $data
            $reg = $data
        }

        Write-Host " ONLINE" -ForegroundColor Green

        $Results += [PSCustomObject]@{
            ComputerName   = $Computer
            OSName         = $os.Caption
            Version        = $os.Version
            Build          = [int]$os.Build
            DisplayVersion = $os.DisplayVersion
            OfflineReason  = 'OK'
        }
    }
    catch {
        Write-Host " OFFLINE" -ForegroundColor Red

        $Results += [PSCustomObject]@{
            ComputerName   = $Computer
            OSName         = 'N/A'
            Version        = 'N/A'
            Build          = 0
            DisplayVersion = 'Offline'
            OfflineReason  = 'UNREACHABLE'
        }
    }
}

#---------------------------------------------CHART DATA------------------------------------------------------------------

$ChartLabels = @(
    'Win11 24H2+',
    'Win11 23H2',
    'Win10 or Older',
    'Offline'
)

$ChartValues = @(
    ($Results | Where-Object { $_.Build -ge 26100 }).Count,
    ($Results | Where-Object { $_.Build -ge 22631 -and $_.Build -lt 26100 }).Count,
    ($Results | Where-Object { $_.Build -gt 0 -and $_.Build -lt 22631 }).Count,
    ($Results | Where-Object { $_.Build -eq 0 }).Count
)

#---------------------------------------------HTML REPORT------------------------------------------------------------------

Write-Host "HTML raporu olusturuluyor..." -ForegroundColor Yellow

New-HTML -TitleText 'Windows Version Inventory' -FilePath $ReportPath {

    New-HTMLSection -HeaderText 'Summary' {
        New-HTMLText -Text "Export Time: $ExportTime"
        New-HTMLText -Text "Total Computers: $($Results.Count)"
    }

    New-HTMLSection -HeaderText 'Build Distribution' {
        New-HTMLChart -Category $ChartLabels {
            New-ChartBar -Name 'Computers' -Value $ChartValues
        }
    }

    New-HTMLSection -HeaderText 'All Computers' {
        New-HTMLTable -DataTable $Results `
            -Title 'Windows Inventory' `
            -PagingLength 25 `
            -Searching `
            -Sorting `
            -ConditionalFormatting {

                New-HTMLTableCondition -Name 'Build' -Operator ge -Value 26100 -BackgroundColor LightGreen
                New-HTMLTableCondition -Name 'Build' -Operator ge -Value 22631 -BackgroundColor Khaki
                New-HTMLTableCondition -Name 'Build' -Operator lt -Value 22631 -BackgroundColor Salmon
            }
    }
}

Write-Host "HTML raporu basariyla olusturuldu: $ReportPath" -ForegroundColor Green
