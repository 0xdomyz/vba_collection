'Turn off and on various settings to speed up macro
Sub turn_offs()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayStatusBar = False
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlCalculationManual
End Sub
Sub turn_ons()
    Application.Calculation = xlCalculationAutomatic
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub    

'Notes
'
'To save status and recover them:
'   Public CalcState As Long
'   CalcState = Application.Calculation
'   Application.Calculation = CalcState
'
'   Public EventState As Boolean
'   EventState = Application.EnableEvents
'   Application.EnableEvents = EventState
'
'   Public PageBreakState As Boolean
'   PageBreakState = ActiveSheet.DisplayPageBreaks
'    ActiveSheet.DisplayPageBreaks = PageBreakState
'
'Other:
'   Range("A1").Select
'   Application.StatusBar = False
'   ActiveSheet.DisplayPageBreaks = False


'Code candidates
dim wb as Workbook
dim a as string
set sht = ThisWorkbook.sheets("Sheet1")
a = sht.range("A1").value
Set wb = Workbooks.Open(a, ReadOnly := True)
wb.Close

