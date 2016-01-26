<#
скрипт запускает экземпляр клиента QlikView выбирает необходимые данные, экспортирует их в excel и отправляет на почту
#>

#Параметры отбора
#Год
$p_year = "2016"
#Месяц
$p_month_namber = 1 #Порядковый номер месяца (янв - 1, дек - 12 )
$p_month = @("янв","Фев","мар","апр","май","июн","июл","авг","сен","окт","ноя","дек")
#День
$p_day = 31

#Функция получения данных с буфера обмена
function Get-Clipboard([switch] $Lines) {
	if($Lines) {
		$cmd = {
			Add-Type -Assembly PresentationCore
			[Windows.Clipboard]::GetText() -replace "`r", '' -split "`n"
		}
	} else {
		$cmd = {
			Add-Type -Assembly PresentationCore
			[Windows.Clipboard]::GetText()
		}
	}
	if([threading.thread]::CurrentThread.GetApartmentState() -eq 'MTA') {
		& powershell -Sta -Command $cmd
	} else {
		& $cmd
	}
}



#Запускаем приложение QlikView
$obj = New-Object -ComObject QlikTech.QlikView

#Задание ожидает запуск приложения QlikView, после чего вводит пароль и нажимет "Ок"
Start-Job –Name QvPasswordIn  -ArgumentList $obj.GetProcessId() –Scriptblock {
	Param($objID)
	#Используя модуль http://wasp.codeplex.com получаем доступ к UI открытого приложения QlikView
	Add-PSSnapin WASP
	out-file -filepath c:\Test\out.txt  -inputobject $objID -Append
	$QvProcess = Get-Process -id $objID #Получаем ссылку на объект запущенног процесса
	$QvMainWindows = Select-Window -InputObject $QvProcess #Основное окно
	while((Select-ChildWindow $QvMainWindows).Title -notmatch "Password")
	{
	}
	Select-ChildWindow $QvMainWindows | Send-Keys "QlikView" | Select-Control | Send-Click
	Select-ChildWindow $QvMainWindows | Select-Control | Send-Click

}

$doc = $obj.OpenDocEx('qvp://optima\QVD_SSA01@192.168.3.76/Application/Stock_Sales_Analysis.qvw',1,"True") 

#Активируем вкладку "Продажи ТТ"
$sheet = $doc.ActivateSheetByID("SH20")


#Переменные параметры
# Выбираем значение Год, Квартал, Месяц
$doc.Fields("Год").Select($p_year)
#$doc.Fields("Квартал").Select("К1")
$doc.Fields("Месяц").Select($p_month[$p_month_namber-1])
# Выбираем "Время отчёта"
$doc.Variables("day").SetContent($p_day,1)

#Получаем таблицу с данными
$ch27 = $doc.GetSheetObject("CH27")

#Сохраняем таблицу в буферобмена и создаём объекты на осонве этой таблицы
$ch27.Copytabletoclipboard("True")

$doc.CloseDoc() #Закрываем документ QlikView
$obj.quit() #Закрываем приложение QlikView

####
# копируем и переименовываем файл шаблона 
####
$filename = "Продажи ТТ $p_day" + $p_month[$p_month_namber-1] + "$p_year"
Copy-Item .\Templates\ПродажиТТ.xlsx .\Tmp\$filename.xlsx
$curDir = $MyInvocation.MyCommand.Definition | split-path -parent
# Create Excel object
$objExcel = new-object -comobject Excel.Application
$objExcel.Visible = $False
$objExcel.displayAlerts = $False
$objWB = $objExcel.Workbooks.Open("$curDir\Tmp\$filename.xlsx")
#Удаляем страницу "data" если есть
try
{
	$objWB.Sheets.Item("data").delete()
} catch
{
}

#Создаём страицу "data"
$objSh_data = $objWB.Sheets.Add()
$objSh_data.Name = "data"
$objSh_data.Activate()
#Вставляем данные с буфера обмена
$objSh_data.paste()
#Сохраняем файл
$objWB.Save()
$objExcel.Quit()




$encoding = [System.Text.Encoding]::UTF8
Send-MailMessage -To "le_it@optima-nv.ru","itnach@optima-nv.ru" -From "QlikView@lamel.biz" -Subject "Отчёт QlikView $filename" -Attachments "$curDir\Tmp\$filename.xlsx" -SmtpServer "192.168.2.17" -Encoding $encoding
#Send-MailMessage -To "le_it@optima-nv.ru","itnach@lamel.biz" -From "QlikView@lamel.biz" -Subject "Report QlikView" -Attachments "c:\test\text.xls" -SmtpServer "192.168.2.17" -Encoding $encoding

#Удаляем файл из папки tmp
Remove-Item "$curDir\Tmp\$filename.xlsx" -Recurse