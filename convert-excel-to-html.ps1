# This script was created to convert Excel charts to html to be viewed in a Browser.
# In my case the files changed on a monthly basis, organized in a fixed folder structure.

While ($true)
{
    $Date = Get-Date -Format "MMMM yyyy"
    $Year = Get-Date -Format "yyyy"
    $Name = "Filename-$Date.xlsx"
    $Excel = New-Object -ComObject "Excel.Application"
    $excel.DisplayAlerts = $false;
    $WorkBook = $Excel.Workbooks.Open("\\share\folder\$Year\$Name")
    $WorkSheet = $WorkBook.Worksheets.Item(1)
    $WorkSheet.SaveAs("\\share\folder\\index.html",44)
    $excel.Quit() # if this doesn't close the process use taskkill /f /im excel.exe
    Start-Sleep 30
}
