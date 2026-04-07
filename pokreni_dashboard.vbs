Set Excel = CreateObject("Excel.Application")
Excel.Visible = False
Excel.DisplayAlerts = False

Set wb = Excel.Workbooks.Open("C:\GITHUB\anketa_web\Anketa_GitHub.xlsm")

Excel.Run "AUTO_index"

wb.Close False
Excel.Quit

Set wb = Nothing
Set Excel = Nothing

Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "C:\GITHUB\anketa_web\upload.bat", 0, True