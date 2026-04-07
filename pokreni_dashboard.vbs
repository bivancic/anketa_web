On Error Resume Next

Dim Excel, wb, WshShell

Set Excel = CreateObject("Excel.Application")
Excel.Visible = False
Excel.DisplayAlerts = False

Set wb = Excel.Workbooks.Open("C:\GITHUB\anketa_web\Anketa_GitHub.xlsm")

' Pokreni makro
Excel.Run "AUTO_index"

' PRIÈEKAJ da Excel završi sve
Do While Excel.Ready = False
    WScript.Sleep 1000
Loop

' Spremi ako treba
wb.Close False

Excel.Quit

' DODATNO osiguranje (ubije ako je ostao)
WScript.Sleep 2000

Set wb = Nothing
Set Excel = Nothing

' Pokreni BAT i ÈEKAJ da završi
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "C:\GITHUB\anketa_web\upload.bat", 0, True

Set WshShell = Nothing