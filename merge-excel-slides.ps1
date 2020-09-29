# Created this script to merge two or more Excel slides into one and export it as HTML.

function Get-WeekNumber1([datetime]$DateTime = (Get-Date)) {
    $cultureInfo = [System.Globalization.CultureInfo]::CurrentCulture
    $cultureInfo.Calendar.GetWeekOfYear($DateTime,$cultureInfo.DateTimeFormat.CalendarWeekRule,$cultureInfo.DateTimeFormat.FirstDayOfWeek)
}
function Get-WeekNumber2([datetime]$DateTime = ((Get-Date).AddDays(7))) {
    $cultureInfo = [System.Globalization.CultureInfo]::CurrentCulture
    $cultureInfo.Calendar.GetWeekOfYear($DateTime,$cultureInfo.DateTimeFormat.CalendarWeekRule,$cultureInfo.DateTimeFormat.FirstDayOfWeek)
}
while($true) {
$Week1 = Get-WeekNumber1
$Week2 = Get-WeekNumber2
$file1 = "\\server-dc\Algemein\KW $Week1.xlsx"
$file2 = "\\server-dc\Algemein\KW $Week2.xlsx"
$Excel = New-Object -ComObject "Excel.Application"
$excel.DisplayAlerts = $false;
$Workbook = $Excel.Workbooks.open($file2)
$Worksheet = $Workbook.WorkSheets.item(“Wochenplan”)
$worksheet.activate() 
$Range = $WorkSheet.Range(“A1:J1”).EntireColumn
$Range.Copy() | out-null
$Workbook = $excel.Workbooks.open($file1)
$Worksheet = $Workbook.Worksheets.item("Wochenplan")
$Range = $Worksheet.Range(“L1”)
$Worksheet.Paste($Range)
$Workbook.SaveAs('\\server-dc\Algemein\Wochenplan Master.html',44) 
$Excel.Quit()
del \\server-dc\Algemein\*.tmp
Start-Sleep –Seconds 60
}
