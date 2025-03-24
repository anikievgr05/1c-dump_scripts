# Функция для загрузки конфигурации из файла
function Load-Config {
    param (
        [string]$configFilePath
    )
    
    $config = @{}
    Get-Content $configFilePath | ForEach-Object {
        # Пропускаем пустые строки и комментарии
        if ($_ -match "^\s*$" -or $_ -match "^\s*#") {
            return
        }
        # Разделяем строку на ключ и значение
        if ($_ -match "^(.*?)\s*=\s*(.*)$") {
            $key = $matches[1].Trim()
            $value = $matches[2].Trim()
            $config[$key] = $value
        }
    }
    return $config
}

# Путь к файлу конфигурации
$configFilePath = ".\config.conf"

# Загружаем конфигурацию
$config = Load-Config -configFilePath $configFilePath

# Инициализация переменных из конфигурации
$server_name = $config["server_name"]
$base_name = ''
$one_c_executable = $config["one_c_executable"]
$disk_name = $config["disk_name"]
$path = $config["folder_path"]
$folder_path = "$disk_name`:\$path"
$bot_token = $config["bot_token"]
$chat_id = $config["chat_id"]
$bases = $config["bases"].Split(",")


# 
# функция для закрытия всех сессий
#
# @return - void
function terminate_all_sessions {

    # Создаем COM-объект подключения к 1С
    $connector = New-Object -ComObject "V83.COMConnector"
    
    try {
        # Подключаемся к агенту сервера 1С
        $AgentConnection = $connector.ConnectAgent($server_address)
        
        # Получаем первый кластер (если их несколько, нужно адаптировать код)
        $Cluster = $AgentConnection.GetClusters()[0]
        
        # Авторизация (пустые логин и пароль, если нет авторизации)
        $AgentConnection.Authenticate($Cluster, $username, $password)
        
        # Получаем список всех сессий для каждой базы
        $sessions = $AgentConnection.GetSessions($Cluster) | Where-Object {
            $_.Infobase.Name -eq $base_name -and
            $_.AppId -ne "SrvrConsole" -and
            $_.AppId -ne "BackgroundJob"
        }
        
        foreach ($session in $sessions) {              
            # Завершаем сессию
            $AgentConnection.TerminateSession($Cluster, $session)
        }
    } catch {
        $errorMessage = $_.Exception.Message
        send_msg -msg "❌ Ошибка при завершении сессий: $errorMessage"
        throw
    }
}

# 
# функция отправляет сообщение в tg got
#
# @param (string) msg - сообщение об отправке
# @return - void
function send_msg {
    param (
        [string]$msg
    )
    
    # url для отправки сообщения
    $url = "https://api.telegram.org/bot$bot_token/sendMessage"
    $utf8Msg = [System.Text.Encoding]::UTF8.GetString([System.Text.Encoding]::GetEncoding("Windows-1251").GetBytes($msg))

    # формируем тело запроса
    $body = @{
        chat_id = $chat_id
        text    = $utf8Msg
    }
    # отправляем POST-запрос
    $response = Invoke-RestMethod -Uri $url -Method Post -Body $body
}

# 
# формироует отчет и отправляет в tg bot
#
# @param (string) start - дата начала выгрузки
# @param (string) end - дата окончания выгрузки
# @param (timespan) time_spent - общее затраченное времмя на выгрузку
# @param (int) file_size - размер файла (бэкапа)
# @param (int) latest_file_size - размер предыдущего файла (бэкапа)
# @return - void
function report {
    param (
        [string]$start,
        [string]$end,
        [timespan]$time_spent,
        [int]$file_size,
        [int]$latest_file_size
    )
    
    # преобразуем время в секунды
    $time_seconds = [math]::Round($time_spent.TotalSeconds, 2)
    
    # получаем дату
    $date = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    
    # получаем разницу размеров файлов
    $difference = 0
    $difference_msg = ""
    if ($file_size -ge $latest_file_size) {
        $difference = $file_size - $latest_file_size
        $difference_in_mb = [math]::Round($difference / 1MB, 2)
        $difference_msg = "новым и последним: $difference_in_mb"
    } elseif ($latest_file_size -ge $file_size) {
        $difference = $latest_file_size - $file_size
        $difference_in_mb = [math]::Round($difference / 1MB, 2)
        $difference_msg = "последним и новым: $difference_in_mb"
    } else {
        $difference_msg = "0"
    }

    # получаем размер в MB
   $file_size_in_mb = [math]::Round($file_size / 1MB, 2)
    # получаем количество совбодного места на диске
    $disk = Get-PSDrive -Name $disk_name
    $free_space = $disk.Free
    $free_space_in_gb = [math]::Round($free_space / 1GB, 2)
    # формируем отчет
    $msg =  @"
===$date===
📢Отчет по #$base_name
Дата начала: $start
Дата окончания: $end
Общее количество затраченного времени в секундах: $time_seconds
Размер файла dt: $file_size_in_mb MB
Разница между $difference_msg MB
Свободного места осталось: $free_space_in_gb GB
"@
    send_msg -msg $msg
}

# 
# функция выгружает информационную базу в dt
#
# @return - void
function unloading_the_information_base {
    #формируем путь к бекапу
    $date = Get-Date -Format "MM_dd_yyyy___HH_mm_ss"
    $full_folder_path = "$folder_path\$base_name"
    if (-not (Test-Path -Path $full_folder_path)) {
        New-Item -ItemType Directory -Path $full_folder_path | Out-Null
    }
    # Путь для сохранения файла .dt
    $output_path = "$full_folder_path\$date.dt"

    # Закрываем все сессии перед выгрузкой
    terminate_all_sessions

    # формируем команду для выгрузки базы
    if ($username -and $password) {
        $command = "CONFIG /DumpIB $output_path /S $server_name\$base_name /N $username /P $password"
    } else {
        $command = "CONFIG /DumpIB $output_path /S $server_name\$base_name"
    }

    # получаем размер последнего файла
    $latest_file = Get-ChildItem -Path $full_folder_path -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $latest_file_size = 0
    # проверяем, найден ли файл
    if ($latest_file) {
        $latest_file_size = $latest_file.Length
    }
    

    # запуск 1С для выполнения выгрузки
    send_msg -msg "🟠 Начинаю выгрузку базы 1С $base_name"
    $start_time = Get-Date
    try {
        Start-Process -FilePath $one_c_executable -ArgumentList $command -Wait -ErrorAction Stop
    } catch {
        send_msg -msg "❌ Ошибка при выполнении выгрузки базы 1С."
        return
    }
    $end_time = Get-Date
    $time_taken = $end_time - $start_time

    # проверка результата
    if (Test-Path $output_path) {
        send_msg -msg "🟢 Выгрузка успешно завершена. Файл сохранен по пути: $output_path"
        # получаем размер нового файла
        $file = Get-Item $output_path
        $file_size = $file.Length
        # отправляем отчет
        report -start      ($start_time.ToString("yyyy-MM-dd HH:mm:ss")) `
               -end        ($end_time.ToString("yyyy-MM-dd HH:mm:ss")) `
               -time_spent $time_taken `
               -file_size  $file_size `
               -latest_file_size  $latest_file_size

    } else {
        send_msg -msg "❌ Ошибка при выгрузке базы 1С."
    }
}
# перебираем базы
foreach ($base in $bases) {
    # Путь к файлу конфигурации
    $configFilePath = ".\settings\$base.conf"
    if (Test-Path $configFilePath) {
        # Загружаем конфигурацию
        $config = Load-Config -configFilePath $configFilePath
        $username = ""
        $password = ""
        if ($config['username'] -and $config['password']) {
            $username = $config['username']
            $password = $config['password']
        }
        $base_name = $base
        unloading_the_information_base
    } else {
        send_msg -msg "❌ Настроек для ИБ $base не существует"
        continue
    }
}