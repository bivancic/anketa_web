On Error Resume Next

Dim Excel, wb, WshShell, oExec

Set Excel = CreateObject("Excel.Application")
Excel.Visible = False
Excel.DisplayAlerts = False

Set wb = Excel.Workbooks.Open("C:\GITHUB\anketa_web\Anketa_GitHub.xlsm")

Excel.Run "AUTO_index"

Do While Excel.Ready = False
    WScript.Sleep 1000
Loop

wb.Close False
Excel.Quit

WScript.Sleep 2000

Set wb = Nothing
Set Excel = Nothing

Set WshShell = CreateObject("WScript.Shell")

Set oExec = WshShell.Exec("cmd /c C:\GITHUB\anketa_web\upload.bat")

Do While oExec.Status = 0
    WScript.Sleep 1000
Loop

Set oExec = Nothing
Set WshShell = Nothing

WScript.Quit