'Return worksheet from named cell
'
'Input: name of the cell, which could be placed at A1 of sheet to act as pivot
Function get_sheet(ByVal ref_nme As String) As Worksheet
    Set get_sheet = Range(ref_nme).Worksheet
End Function

'Turn off and on various settings to speed up macro
Sub turn_offs()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = False
    Application.Calculation = xlCalculationManual
End Sub
Sub turn_ons()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub    
