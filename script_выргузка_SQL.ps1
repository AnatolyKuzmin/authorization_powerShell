# задать даты
#$date = (Get-Date).AddDays(-1).ToString("dd.MM.yyyy")
$m = (get-date).AddDays(-1).ToString("MM")
$y = (get-date).AddDays(-1).ToString("yyy")
$d = (get-date).AddDays(-1).ToString("dd")

# сохранение лога
$ListTo = Get-Content "\\путь\папкаОтчета\logs.txt" -Encoding UTF8

# ссылка на сохранение файла
$path = "\\путь\папкаОтчета\"
$files = "$m.$y"+"_отчет.csv"
$export = $path + $files

# если потребуется фильтрация############################################################################
$path_queues = "\\путь\папкаОтчета\Фильтры.xlsx"

$excel = New-Object -comobject Excel.Application
#$excel.visible = $true
$excel.DisplayAlerts = $false
$password = ""
$workbook = $excel.Workbooks.open($path_queues,0,1,5,$password,$password)
$worksheet = $workbook.Worksheets.Item("названиеЛиста")
$queues = $worksheet.Range(“A2”).Text

$workbook.close($false)
$excel.Quit()
(Get-Process -Name "excel").Kill()
Remove-Variable -Name excel
#########################################################################################################

# Параметры подключения
$dbServer = "12.34.567.89" 
$dbName = "название"
$dbUser = "пользователь"
$dbPass = "пароль"
$port = "1234"
[string]$compConStr = "Driver={PostgreSQL UNICODE};Server=$dbServer;Port=$port;Database=$dbName;Uid=$dbUser;Pwd=$dbPass;" 
$connectlog = ""

# Создание и открытие соединения
$cnDB= New-Object System.Data.Odbc.OdbcConnection($compConStr)
try {
    $cnDB.Open()
    Write-Host "Успешное подключение к PSQL..." -BackgroundColor Cyan
} catch {
    Write-Host "Ошибка в подключении к PSQL..." -BackgroundColor Cyan
    $connectlog = "Ошибка в подключении к PSQL."
    #return
}

# SQL
$sql =
@"
SELECT *
FROM таблица
"@


# Создание команды для выполнения SQL-запроса
$command = New-Object System.Data.Odbc.OdbcCommand($sql,$cnDB)
$command.CommandTimeout = 0

# Создание DataSet для хранения результатов
$ds = New-Object system.Data.DataSet
$da = New-Object system.Data.odbc.odbcDataAdapter($command)

# Заполнение DataSet результатами SQL-запроса
[void]$da.fill($ds)

# Проверка на пустой результат и экспорт в CSV
if ($ds.Tables[0].Rows.Count -eq 0) {
    # Если результат пуст, создание и экспорт CSV-файла только с заголовком
    $header = ""
    foreach ($col in $ds.Tables[0].Columns) {
        $header += $col.ColumnName +";"
    }
    $header.Remove($header.Length-1,1) | Out-File $export -Encoding UTF8
} else {
    # Если есть результат, экспорт данных в CSV-файл
    $ds.Tables[0] | Export-Csv $export -NoTypeInformation -Encoding UTF8 -Delimiter ';'
}

###################Таблица для письма###################
# Расчет количества уникальных значений по датам
$CONNID_count =  $ds.Tables[0] | Group-Object DAY | Select Name, Count | Sort Name

# Построение HTML-таблицы для письма
$html = "<style>table, th, td {border: 1px solid;text-align: center;}</style><table><tr><td>Дата</td><td>Кол-во строк</td></tr>"
if ( $CONNID_count.Count -eq 0 ) {
    # При условии что SQL пуст
    $html += "<tr><td>" + "$d.$m.$y" + "</td><td>" + $CONNID_count.Count + "</td></tr>"
} elseif ( !$CONNID_count[1] ) {
    # При условии что в SQL одна строка
    $html += "<tr><td>" + $CONNID_count[0].Name + "</td><td>" + $CONNID_count[0].Count + "</td></tr>"
} else {
    # При условии что в SQL много строк
    for($i = 0; $i -le $CONNID_count.Length-1; $i++) {
        $html += "<tr><td>" + $CONNID_count[$i].Name + "</td><td>" + $CONNID_count[$i].Count + "</td></tr>"
    }
}
$html += "</table>"

###################Проверка файла#######################
# Получение времени последней модификации экспортированного файла
$datetime = (Get-ChildItem -Path $export | Select LastWriteTime).LastWriteTime

# Проверка существования файла
$test_path = Test-Path -Path $export

###########Создание и заполнение письма#################
# Создание объектов Outlook для отправки электронного письма
$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)

$link = "\\папка\Scripts"
# Проверка наличия файла и даты последней модификации
if ($test_path -ne $false -and ($datetime.ToString("dd.MM.yyyy") -eq (get-date).ToString("dd.MM.yyyy"))){
    # Если файл существует и был изменен сегодня, формирование письма с результатами
    $Mail.HTMLBody ="Добрый день!<BR>$connectlog<BR>$($html)"
}else{
    # Если файл не существует или не изменен сегодня, формирование письма об ошибке
    $Mail.HTMLBody ="Добрый день!<BR><BR>Ошибка при выгрузке.<BR>$connectlog<BR>Скрипт - <a href=$link>postgresql_bridge_chat_details.ps1</a><BR><BR>$Error"
}

# Установка отправителя, получателей и отправка письма
$Mail.Subject = "Лог - SQL"
$Mail.SentOnBehalfOfName = "отправитель@почта.ru"
$Mail.To = $ListTo
$Mail.Send()
#$Mail.Display()

# Очистка ресурсов и объектов Outlook
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Outlook) | Out-Null
$cnDB.Close()