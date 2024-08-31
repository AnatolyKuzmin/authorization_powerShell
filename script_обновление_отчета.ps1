# закрыть процесс excel
(Get-Process -Name "excel").Kill()

# задать даты
#$date = (Get-Date).AddDays(-1).ToString("dd.MM.yyyy")
$yyyy = $((Get-Date).AddDays(-1).ToString("yyyy"))
$m = $((Get-Date).AddDays(-1).ToString("MMMM"))
$mm = $((Get-Date).AddDays(-1).ToString("MM"))

# задаем путь
$subPath = "\\путь\папкаОтчета\$yyyy"
$path = "\\путь\папкаОтчета\$yyyy\$yyyy"+"_отчет.xlsx"

# Создание файла "отчета" при его отсутствии
if (!(Test-Path $path)){
	$yyyyPastMont = $((Get-Date).AddYears(-1).ToString("yyyy"))
	$pathPast = "\\путь\папкаОтчета\$yyyyPastMont\$yyyyPastMont"+"_отчет.xlsx"
    Copy-Item -Path $pathPast -Destination $path
}

# Обновление отчёта "отчета"
$excel = New-Object -comobject Excel.Application
$excel.visible = $true
$excel.DisplayAlerts = $false
$password = ""
$updatelinks = 0
$workbook = $excel.Workbooks.open($path,0,0,5,$password,$password)
$xlCalculationManual = -4135
$xlCalculationAutomatic = -4105
$excel.Calculation = $xlCalculationAutomatic
$excel.CalculateFullRebuild() | Out-Null
$conn = $workbook.Connections

$conn | ForEach-Object {
    write-host $_.Name
    $_.Refresh()
    while($_.OLEDBConnection.Refreshing -or !$excel.Ready){
        Start-Sleep -s 1
    }
}
Start-Sleep -s 25

# принудительно выполнить полное вычисление данных и перестроить зависимости.
$excel.CalculateFullRebuild() | Out-Null

# Какая нибудь Проверка
$worksheets = $workbook.WorkSheets.item("ЛистПроверка")
$cell = $worksheets.range("A1").Text
# Адресаты из excel
$ToOk = $worksheets.range("A2").Text
$ToErr = $worksheets.range("A3").Text

# Формирование письма
if ($cell -eq 1 -and $cellBI -eq 1)
{
$To = $ToOk
$link = "ПУТЬ"

$Body = @"
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
Добрый день!
<br><br>
Коллеги, обновленный отчет доступен по <a href= "$link">ссылке</a>.<br>
<br><br>
Контактные лица:<br>
(Имя Фамилия e-mail: почта@ответсвенного, mob: +790000000000)<br>
"@
}
else
{
$To = $ToErr
$Body = @"
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
Добрый день
!<br><br>
Ошибка при обновлении.
"@
}

#Отправка письма
$outlook = New-Object -ComObject Outlook.Application
$mail = $Outlook.CreateItem(0)
$mail.SentOnBehalfOfName = "почта@отправителя"
$mail.To = $To
#$mail.CC = $CC
$mail.Subject = "Название отчета $m"
$mail.HTMLBody = $Body

#$mail.Display()
$mail.Send()

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
(Get-Process -Name "excel").Kill()
Remove-Variable -Name excel
