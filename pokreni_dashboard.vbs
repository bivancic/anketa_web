Option Explicit
On Error Resume Next

Dim Excel, wb, WshShell, oExec

' Launch Excel invisibly
Set Excel = CreateObject("Excel.Application")
Excel.Visible = False
Excel.DisplayAlerts = False

' Open the workbook
Set wb = Excel.Workbooks.Open("C:\GITHUB\anketa_web\Anketa_GitHub.xlsm")

If Err.Number <> 0 Then
    WScript.Echo "Error opening workbook: " & Err.Description
    Excel.Quit
    WScript.Quit 1
End If

' Run the macro
Excel.Run "AUTO_index"

If Err.Number <> 0 Then
    WScript.Echo "Error running macro: " & Err.Description
    wb.Close False
    Excel.Quit
    WScript.Quit 1
End If

' Wait for Excel to finish
Do While Excel.Ready = False
    WScript.Sleep 1000
Loop

' Close workbook without saving, quit Excel
wb.Close False
Excel.Quit

Set wb = Nothing
Set Excel = Nothing

WScript.Sleep 2000

' Run the batch file
Set WshShell = CreateObject("WScript.Shell")
Set oExec = WshShell.Exec("cmd /c C:\GITHUB\anketa_web\upload.bat")

Do While oExec.Status = 0
    WScript.Sleep 1000
Loop

Set oExec = Nothing
Set WshShell = Nothing

WScript.Quit 0