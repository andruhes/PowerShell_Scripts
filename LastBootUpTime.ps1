# Список виртуальных машин, которые нужно исключить
$ExcludedVMs = @(
    'Bitrix1',
    'Bitrix2',
    'crm'
)

# Укажите имена хостов
$hosts = @("host1", "host2", "host3")

# Укажите параметры для отправки email
$smtpServer = "192.168.1.5"
$toEmail = "it@andruhes.ru"
$fromEmail = "host1-powershell@andruhes.ru" # Замените на ваш email, если необходимо
$subject = "Win VM uptime - Cluster 1"

# Функция для получения информации о последней загрузке и реальном времени работы
function Get-VMLastBootInfo {
    param(
        [Parameter(Mandatory=$true)]
        [string]$HostName,
        [Parameter(Mandatory=$true)]
        [string]$VMName
    )
    
    try {
        # Получаем объект виртуальной машины
        $vm = Get-VM -Name $VMName -ComputerName $HostName -ErrorAction Stop
        
        # Получаем информацию о времени последнего запуска через WMI
        $lastBootUpTime = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $vm.Name |
                          Select-Object -ExpandProperty LastBootUpTime
        
        # Преобразуем время в нужный формат
        if ($lastBootUpTime) {
            $bootTime = [Management.ManagementDateTimeConverter]::ToDateTime($lastBootUpTime)
            $formattedDate = $bootTime.ToString("dd.MM.yy")  # Формат даты: 29.01.25
            $formattedTime = $bootTime.ToString("HH:mm")     # Формат времени: 10:22
            $lastBootUpTimeFormatted = "$formattedDate $formattedTime"
            
            # Рассчитываем реальный uptime
            $uptime = New-TimeSpan -Start $bootTime -End (Get-Date)
            $realUptime = "{0}d {1}h {2}m" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
        } else {
            $lastBootUpTimeFormatted = "Not available"
            $realUptime = ""
        }
        
        # Возвращаем результат в виде хэш-таблицы
        @{
            Host = $HostName
            VMName = $VMName
            LastBootTime = $lastBootUpTimeFormatted
            RealUptime = $realUptime
        }
    } catch {
        # Если произошла ошибка, возвращаем сообщение об ошибке
        @{
            Host = $HostName
            VMName = $VMName
            LastBootTime = "Error: $_"
            RealUptime = ""
        }
    }
}

# Перебираем все хосты и собираем информацию о виртуальных машинах
$results = foreach ($currentHost in $hosts) {
    try {
        # Получаем список виртуальных машин
        $vmList = Get-VM -ComputerName $currentHost | Where-Object { $ExcludedVMs -notcontains $_.Name }
        
        foreach ($vm in $vmList) {
            # Вызываем функцию для получения информации о последней загрузке и реальном времени работы
            Get-VMLastBootInfo -HostName $currentHost -VMName $vm.Name
        }
    } catch {
        # Если произошла ошибка при получении списка виртуальных машин, добавляем сообщение об ошибке
        @{
            Host = $currentHost
            VMName = "N/A"
            LastBootTime = "Error: $_"
            RealUptime = ""
        }
    }
}

# Формируем HTML-таблицу
$html = "<style>
table {
    border-collapse: collapse;
    width: 100%;
}
th, td {
    text-align: left;
    padding: 8px;
    border-bottom: 1px solid #ddd;
}
tr:nth-child(even){background-color: #f2f2f2;}
</style>"

$html += "<h2>Win VM uptime - Cluster 1</h2><br/>"

# Первая таблица с фильтром по uptime > 1 мин и < 24 часа, исключая 'Not available' и пустые значения
$html += "<h2>Win VM uptime < 24 h</h2><br/>"
$html += "<table>"
$html += "<tr><th>Host</th><th>VM Name</th><th>Last Boot Time</th><th>Real Uptime</th></tr>"

# Фильтруем результаты, оставляя только те, у которых uptime больше 1 минуты и меньше 24 часов
$lessThan24Hours = $results | Where-Object { 
    $timeParts = $_.RealUptime.Split(' ')
    $days = [int]$timeParts[0].Replace('d', '')
    $hours = [int]$timeParts[1].Replace('h', '')
    $minutes = [int]$timeParts[2].Replace('m', '')
    $totalMinutes = ($days * 1440) + ($hours * 60) + $minutes
    $totalMinutes -gt 1 -and $totalMinutes -lt 1440 -and $_.RealUptime -ne '' 
} | Sort-Object { 
    $timeParts = $_.RealUptime.Split(' ')
    $days = [int]$timeParts[0].Replace('d', '')
    $hours = [int]$timeParts[1].Replace('h', '')
    $minutes = [int]$timeParts[2].Replace('m', '')
    $totalMinutes = ($days * 1440) + ($hours * 60) + $minutes
    $totalMinutes
}

foreach ($result in $lessThan24Hours) {
    $html += "<tr>"
    $html += "<td>$($result.Host)</td>"
    $html += "<td>$($result.VMName)</td>"
    $html += "<td>$($result.LastBootTime)</td>"
    $html += "<td>$($result.RealUptime)</td>"
    $html += "</tr>"
}

$html += "</table>"

# Вторая таблица со всеми ВМ
$html += "<h2>All Win VMs</h2><br/>"
$html += "<table>"
$html += "<tr><th>Host</th><th>VM Name</th><th>Last Boot Time</th><th>Real Uptime</th></tr>"

foreach ($result in $results) {
    $html += "<tr>"
    $html += "<td>$($result.Host)</td>"
    $html += "<td>$($result.VMName)</td>"
    $html += "<td>$($result.LastBootTime)</td>"
    $html += "<td>$($result.RealUptime)</td>"
    $html += "</tr>"
}

$html += "</table>"

# Отправляем отчёт по электронной почте
try {
    Send-MailMessage -SmtpServer $smtpServer -To $toEmail -From $fromEmail -Subject $subject -BodyAsHtml $html
    Write-Host "Отчёт успешно отправлен."
} catch {
    Write-Host "Ошибка при отправке отчёта: $_"
}