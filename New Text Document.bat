set ExcelFilePath="E:\001_QA_Generator\snap_tool\convert.xlsx"
set SheetsToKeep="Conversions", "part_number"

:: Create a PowerShell command to delete sheets except for specified ones
set PowerShellCommand=^
    $Excel = New-Object -ComObject Excel.Application ^|^|^
    $Workbook = $Excel.Workbooks.Open('%ExcelFilePath%') ^|^|^
    $Sheets = $Workbook.Sheets ^|^|^
    $SheetsToDelete = $Sheets | Where-Object { $_.Name -notin %SheetsToKeep% } ^|^|^
    $SheetsToDelete | ForEach-Object { $_.Delete() } ^|^|^
    $Workbook.Save() ^|^|^
    $Excel.Quit() ^|^|^
    Remove-Variable -Name Excel -Force -ErrorAction SilentlyContinue

:: Run the PowerShell command
powershell -Command "%PowerShellCommand%"