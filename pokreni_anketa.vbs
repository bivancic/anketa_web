On Error Resume Next

Dim Excel, wb

Set Excel = CreateObject("Excel.Application")
Excel.Visible = False
Excel.DisplayAlerts = False

Set wb = Excel.Workbooks.Open("C:\GITHUB\anketa_web\Anketa_GitHub.xlsm")

' pokreni makro
Excel.Run "AUTO_index"

' èekaj da završi
Do While Excel.Ready = False
    WScript.Sleep 1000
Loop

wb.Close False
Excel.Quit

Set wb = Nothing
Set Excel = Nothing

WScript.Quit