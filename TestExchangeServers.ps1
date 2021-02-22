# автор Дмитрий Карпунин
# C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit -command ". 'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell; C:\Scripts\TestExchangeServers_v3.ps1 "

$ErrorActionPreference="SilentlyContinue"

#---------------------------------------------------------------------------------------------

$Title = "Exchange"

$MessageFrom = "Exchange@Powershell"

$MessageTo = "exchangeadmin@company.ru"

$MessageSMTPServer = "<smtp_server_address>"

$MessageSubject = "TestExchangeServersReport от " + (Get-Date -Format "dd.MM.yyyy")

$HTMLStart = " `
    <html> `
    <head> `
    </head> `
    <body style='font-family:Geneva, Arial, Helvetica, sans-serif;font-size:0.8em'> `
    <h1 style='font-size:1.3em'>Отчёт о состоянии серверов</h1> `
"

$MessageBody = $HTMLStart

$TestDisable = "<div style='color:Red'>Тест не пройден</div>"

$TableTop = " `
            <table style='border-collapse:collapse;border-bottom:1px solid #696969'> `
            <tr style='background-color:black;color:white;padding:5px;font-family:Geneva, Arial, Helvetica, sans-serif;font-size:0.8em'> `
            "

$TRTop = "<tr style='background-color:#CCC;padding:5px;font-family:Geneva, Arial, Helvetica, sans-serif;font-size:0.8em'>"

$TRBot = "</tr>"

$TableBot += "</table>"

#---------------------------------------------------------------------------------------------

$MessageBody += "<h3 style='font-size:1.1em'>$Title</h3>"

$Servers = Get-ExchangeServer

foreach ($Server in $Servers) {

    $ServerName = $Server.Name

    $MessageBody += "<hr><h2 style='font-size:1.2em'>$ServerName</h2>"

    $MessageBody += "<h3 style='font-size:1.1em'>Состояние сервисов</h3>"

    $ServicesHealth = $Server | Test-ServiceHealth

    foreach ($ServiceHealth in $ServicesHealth) {

        $Role = $ServiceHealth.Role

        $MessageBody += "<div>$Role</div>"

        $ServicesRunning = $ServiceHealth.ServicesRunning

        $MessageBody += "<div><ul style='list-style:square'>"

        foreach ($ServiceRunning in $ServicesRunning) {

            $MessageBody += "<li style='color:Green'>$ServiceRunning</li>"

        }

        $ServicesNotRunning = $ServiceHealth.ServicesNotRunning

        foreach ($ServiceNotRunning in $ServicesNotRunning) {

            $MessageBody += "<li style='color:Red'><strong>$ServiceNotRunning</strong></li>"

        }

        $MessageBody += "</ul></div>"

    }

    $MessageBody += "<h3 style='font-size:1.1em'>Состояние дисков</h3>"

    $MessageBody += $TableTop
    $MessageBody += "<td>Диск</td>"
    $MessageBody += "<td>Размер (ГБ)</td>"
    $MessageBody += "<td>Свободно (ГБ)</td>"
    $MessageBody += "<td>Использовано (ГБ)</td>"
    $MessageBody += "<td>Свободно %</td>"
    $MessageBody += $TRBot

    $Disks = Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" -ComputerName $ServerName

    if (-not $Disks) {
            
        $ND = "<span style='color:red'>н.д.</span>"

        $MessageBody += $TRTop
        $MessageBody += "<td style='border-bottom:1px solid #696969'>$ND</td>"
        $MessageBody += "<td style='border-bottom:1px solid #696969'>$ND</td>"
        $MessageBody += "<td style='border-bottom:1px solid #696969'>$ND</td>"
        $MessageBody += "<td style='border-bottom:1px solid #696969'>$ND</td>"
        $MessageBody += "<td style='border-bottom:1px solid #696969'>$ND</td>"
                
    } else {

        foreach ($Disk in $Disks) {

            $DeviceID = $Disk.DeviceID
            $Size = "{0:N1}" -f ($Disk.Size / 1gb)
            $Free = "{0:N1}" -f ($Disk.Freespace / 1gb)
            $Used = "{0:N1}" -f (($Disk.Size-$Disk.Freespace) / 1gb)
            $PerFree = "{0:P0}" -f ($Disk.Freespace / $Disk.Size)
            $PerFreeVal = [math]::Round($Disk.Freespace / $Disk.Size * 100)

            $MessageBody += $TRTop
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$DeviceID</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Size</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Free</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Used</td>"

            if ($PerFreeVal -le 5) {
              
                $MessageBody += "<td style='background-color:red;color:white;border-bottom:1px solid #696969'>$PerFree</td>"
            
            } else {

                if ($PerFreeVal -le 10 -and $PerFreeVal -gt 5) {

                    $MessageBody += "<td style='background-color:yellow;border-bottom:1px solid #696969'>$PerFree</td>"

                } else {

                    $MessageBody += "<td style='border-bottom:1px solid #696969'>$PerFree</td>"

                }

            }

            $MessageBody += $TRBot

        }

    }

    $MessageBody += $TableBot

    $MessageBody += "<h3 style='font-size:1.1em'>Подключение к ящикам</h3>"

    $Result = Test-MAPIConnectivity -Server $ServerName

    if ($Result) {

        $MessageBody += $TableTop
        $MessageBody += "<td>База данных</td>"
        $MessageBody += "<td>Результат</td>"
        $MessageBody += $TRBot

        $Nodes = $Result

        foreach ($Node in $Nodes) {

            $Database = $Node.Database
            $Result = $Node.Result

            $MessageBody += $TRTop
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Database</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Result</td>"
            $MessageBody += $TRBot

        }

        $MessageBody += $TableBot

    } else {

        $MessageBody += $TestDisable

    }

    $MessageBody += "<h3 style='font-size:1.1em'>Успешная отправка/доставка почты из системного почтового ящика</h3>"
    
    $Result = Test-Mailflow -Identity $ServerName

    if ($Result) {

        $MessageBody += $TableTop
        $MessageBody += "<td>Результат отправки</td>"
        $MessageBody += "<td>Время задержки</td>"
        $MessageBody += "<td>Успешность</td>"
        $MessageBody += $TRBot

        $Nodes = $Result

        foreach ($Node in $Nodes) {

            $TestMailflowResult = $Node.TestMailflowResult
            $MessageLatencyTime = $Node.MessageLatencyTime
            $IsValid = $Node.IsValid

            $MessageBody += $TRTop
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$TestMailflowResult</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$MessageLatencyTime</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$IsValid</td>"
            $MessageBody += $TRBot

        }

        $MessageBody += $TableBot
    
        
    } else {

        $MessageBody += $TestDisable
    
    }

    $MessageBody += "<h3 style='font-size:1.1em'>Состояние копий почтовых баз</h3>"

    $Result = Get-MailboxDatabaseCopyStatus -Server $ServerName

    if ($Result) {

        $MessageBody += $TableTop
        $MessageBody += "<td>Имя</td>"
        $MessageBody += "<td>Статус</td>"
        $MessageBody += "<td>Длина очереди копирования</td>"
        $MessageBody += "<td>Длина очереди воспроизведения</td>"
        $MessageBody += "<td>Время последней проверки</td>"
        $MessageBody += "<td>Состояние индекса контента</td>"
        $MessageBody += $TRBot

        $Nodes = $Result

        foreach ($Node in $Nodes) {

            $Name = $Node.Name
            $Status = $Node.Status
            $CopyQueueLength = $Node.CopyQueueLength
            $ReplayQueueLength = $Node.ReplayQueueLength
            $LastInspectedLogTime = $Node.LastInspectedLogTime
            $ContentIndexState = $Node.ContentIndexState

            $MessageBody += $TRTop
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Name</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Status</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$CopyQueueLength</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$ReplayQueueLength</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$LastInspectedLogTime</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$ContentIndexState</td>"
            $MessageBody += $TRBot

        }

        $MessageBody += $TableBot

    } else {

        $MessageBody += $TestDisable

    }

    $MessageBody += "<h3 style='font-size:1.1em'>Служба Autodiscover</h3>"

    $Result = Test-OutlookWebServices -ClientAccessServer $ServerName

    if ($Result) {

        $MessageBody += $TableTop
        $MessageBody += "<td>Источник</td>"
        $MessageBody += "<td>Получатель</td>"
        $MessageBody += "<td>Сценарий</td>"
        $MessageBody += "<td>Результат</td>"
        $MessageBody += "<td>Задержка (мс)</td>"
        $MessageBody += $TRBot

        $Nodes = $Result

        foreach ($Node in $Nodes) {

            $Source = $Node.Source
            $ServiceEndpoint = $Node.ServiceEndpoint
            $Scenario = $Node.Scenario
            $Result = $Node.Result
            $Latency = $Node.Latency

            $MessageBody += $TRTop
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Source</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$ServiceEndpoint</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Scenario</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Result</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Latency</td>"
            $MessageBody += $TRBot

        }

        $MessageBody += $TableBot

    } else {

        $MessageBody += $TestDisable

    }

    $MessageBody += "<h3 style='font-size:1.1em'>Работоспособность Outlook Web App</h3>"

    $Result = Test-OwaConnectivity -ClientAccessServer $ServerName

    if ($Result) {

        $MessageBody += $TableTop
        $MessageBody += "<td>Сервер</td>"
        $MessageBody += "<td>Сайт</td>"
        $MessageBody += "<td>Сценарий</td>"
        $MessageBody += "<td>Результат</td>"
        $MessageBody += "<td>Задержка (мс)</td>"
        $MessageBody += $TRBot

        $Nodes = $Result

        foreach ($Node in $Nodes) {

            $CasServer = $Node.ClientAccessServerShortName
            $LocalSite = $Node.LocalSite
            $Scenario = $Node.Scenario
            $Result = $Node.Result
            $Latency = $Node.Latency

            $MessageBody += $TRTop
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$CasServer</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$LocalSite</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Scenario</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Result</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Latency</td>"
            $MessageBody += $TRBot

        }

        $MessageBody += $TableBot

    } else {

        $MessageBody += $TestDisable

    }

    $MessageBody += "<h3 style='font-size:1.1em'>Функциональность веб-служб Exchange</h3>"

    $Result = Test-WebServicesConnectivity -ClientAccessServer $ServerName

    if ($Result) {

        $MessageBody += $TableTop
        $MessageBody += "<td>Источник</td>"
        $MessageBody += "<td>Получатель</td>"
        $MessageBody += "<td>Сценарий</td>"
        $MessageBody += "<td>Результат</td>"
        $MessageBody += "<td>Задержка (мс)</td>"
        $MessageBody += $TRBot

        $Nodes = $Result

        foreach ($Node in $Nodes) {

            $Source = $Node.Source
            $ServiceEndpoint = $Node.ServiceEndpoint
            $Scenario = $Node.Scenario
            $Result = $Node.Result
            $Latency = $Node.Latency

            $MessageBody += $TRTop
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Source</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$ServiceEndpoint</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Scenario</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Result</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Latency</td>"
            $MessageBody += $TRBot

        }

        $MessageBody += $TableBot

    } else {

        $MessageBody += $TestDisable

    }

    $MessageBody += "<h3 style='font-size:1.1em'>Служба репликации почтовых ящиков</h3>"

    $Result = Test-MRSHealth -Identity $ServerName

    if ($Result) {

        $MessageBody += $TableTop
        $MessageBody += "<td>Проверка</td>"
        $MessageBody += "<td>Пройдена</td>"
        $MessageBody += "<td>Сообщение</td>"
        $MessageBody += $TRBot

        $Nodes = $Result

        foreach ($Node in $Nodes) {

            $Check = $Node.Check
            $Passed = $Node.Passed
            $Message = $Node.Message

            $MessageBody += $TRTop
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Check</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Passed</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Message</td>"
            $MessageBody += $TRBot

        }

        $MessageBody += $TableBot

    } else {

        $MessageBody += $TestDisable

    }

    $MessageBody += "<h3 style='font-size:1.1em'>Состояние очереди сообщений</h3>"

    $Result = Get-Queue -Server $ServerName

    if ($Result) {

        $MessageBody += $TableTop
        $MessageBody += "<td>Источник</td>"
        $MessageBody += "<td>Тип доставки</td>"
        $MessageBody += "<td>Статус</td>"
        $MessageBody += "<td>Количество сообщений</td>"
        $MessageBody += "<td>Скорость</td>"
        $MessageBody += "<td>Уровень риска</td>"
        $MessageBody += "<td>Исходящий IP-пул</td>"
        $MessageBody += "<td>Следующий узел</td>"
        $MessageBody += $TRBot

        $Nodes = $Result

        foreach ($Node in $Nodes) {

            $Identity = $Node.Identity
            $DeliveryType = $Node.DeliveryType
            $Status = $Node.Status
            $MessageCount = $Node.MessageCount
            $Velocity = $Node.Velocity
            $RiskLevel = $Node.RiskLevel
            $OutboundIPPool = $Node.OutboundIPPool
            $NextHopDomain = $Node.NextHopDomain

            $MessageBody += $TRTop
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Identity</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$DeliveryType</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Status</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$MessageCount</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Velocity</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$RiskLevel</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$OutboundIPPool</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$NextHopDomain</td>"
            $MessageBody += $TRBot

        }

        $MessageBody += $TableBot

    } else {

        $MessageBody += $TestDisable

    }

    $MessageBody += "<h3 style='font-size:1.1em'>Сертификаты Exchange</h3>"

    $Result = Get-ExchangeCertificate -Server $ServerName

    if ($Result) {

        $MessageBody += $TableTop
        $MessageBody += "<td>Издатель</td>"
        $MessageBody += "<td>Самоподписной</td>"
        $MessageBody += "<td>Действителен до</td>"
        $MessageBody += "<td>Сервисы</td>"
        $MessageBody += "<td>Статус</td>"
        $MessageBody += $TRBot

        $Nodes = $Result

        foreach ($Node in $Nodes) {

            $Issuer = $Node.Issuer
            $IsSelfSigned = $Node.IsSelfSigned
            $NotAfter = $Node.NotAfter
            $Services = $Node.Services
            $Status = $Node.Status

            $MessageBody += $TRTop
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Issuer</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$IsSelfSigned</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$NotAfter</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Services</td>"
            $MessageBody += "<td style='border-bottom:1px solid #696969'>$Status</td>"
            $MessageBody += $TRBot

        }

        $MessageBody += $TableBot

    } else {

        $MessageBody += $TestDisable

    }

}

$MessageBody += "<h3 style='font-size:1.1em'>Доступность DAG</h3>"

$DAGs = Get-DatabaseAvailabilityGroup

if ($DAGs) {

    foreach ($DAG in $DAGs) {

        $MessageBody += $TableTop
        $MessageBody += "<td>Имя</td>"
        $MessageBody += "<td>Серверы</td>"
        $MessageBody += "<td>Доступен</td>"
        $MessageBody += $TRBot

        $Name = $DAG.Name
        $Servers = ($DAG.Servers -split (" ")) -join ", "
        $IsValid = $DAG.IsValid

        $MessageBody += $TRTop
        $MessageBody += "<td style='border-bottom:1px solid #696969'>$Name</td>"
        $MessageBody += "<td style='border-bottom:1px solid #696969'>$Servers</td>"
        $MessageBody += "<td style='border-bottom:1px solid #696969'>$IsValid</td>"
        $MessageBody += $TRBot
        $MessageBody += $TableBot

    }

}

$ReportDateTime = Get-Date -Format "dd.MM.yyyy HH:mm"

$HostName = $env:COMPUTERNAME

$MessageBody += "<hr><p style='text-align:right'>Отчёт создан $ReportDateTime ($HostName)</p>"

$HTMLEnd = " `
    </body> `
    </html> `
"

$MessageBody += $HTMLEnd

Send-MailMessage -Body $MessageBody -BodyAsHtml -From $MessageFrom -To $MessageTo -SmtpServer $MessageSMTPServer -Subject $MessageSubject -Encoding Unicode

$OutFile = "$PSScriptRoot\ExchangeServersReport.html"

$MessageBody | Out-File -FilePath $OutFile -Encoding utf8 -Force