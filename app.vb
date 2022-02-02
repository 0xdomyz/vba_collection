'Application level code snippets

Application.ScreenUpdating = False
Application.CutCopyMode = False
Application.Calculation = xlCalculationManual
Application.Calculation = xlCalculationAutomatic
Application.StatusBar = False

Sub
End Sub

dim wb as Workbook

dim a as string

set sht = ThisWorkbook.sheets("Sheet1")

a = sht.range("A1").value

Set wb = Workbooks.Open(a, ReadOnly := True)

wb.Close

