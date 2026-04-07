Option Explicit
On Error Resume Next

Dim Excel, wb, WshShell, exitCode

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

' Run the batch file - hidden window (0), wait for it to finish (True)
Set WshShell = CreateObject("WScript.Shell")
exitCode = WshShell.Run("cmd /c C:\GITHUB\anketa_web\upload.bat", 0, True)

Set WshShell = Nothing

WScript.Quit exitCode