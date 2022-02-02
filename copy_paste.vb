'Copy paste a block of cells
'
'Copy a continuous row, column, or table of cells.
'Continuous means no empty cells on top and left edges of the block.
'Then paste into another location represented by a cell. Top left corner of the 
'block is to be placed into this representative cell.
'
'Inputs:
'   tl_orig : Top left cell of the original block.
'   rep_cel : Representative cell of the target location.
'   pst_mod : Optional Excel paste mode string.
'       Either : "value", "value_format", "formula", "all"
'       Default : "value"
'
'Effects:
'   A copy paste action.
'
'Notes on Excel paste modes:
'   https://docs.microsoft.com/en-us/office/vba/api/Excel.XlPasteType
'
'Example
'
'Copy paste a block of cells from A1 to E1
'   A  B
'1  1  2
'2  3  4
'Call cp_blk(Range("A1"), Range("E1"))
'   E  F
'1  1  2
'2  3  4
Sub cp_blk( _
    ByRef tl_orig As Range, _
    ByRef rep_cel As Range, _
    Optional ByVal pst_mod As String = "value")

Dim br As Range
Dim cold As Integer
Dim rowd As Integer
cold = sf_end(tl_orig, "right").Column - tl_orig.Column
rowd = sf_end(tl_orig, "down").Row - tl_orig.Row
Set br = tl_orig.Offset(rowd, cold)

Range(tl_orig, br).Copy
If pst_mod = "value" Then
    rep_cel.PasteSpecial Paste:=xlPasteValues, _
    Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Elseif pst_mod = "value_format" Then
    rep_cel.PasteSpecial Paste:=xlPasteValuesAndNumberFormats, _
    Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Elseif pst_mod = "formula" Then
    rep_cel.PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, _
    Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Elseif pst_mod = "all" Then
    rep_cel.PasteSpecial Paste:=xlPasteAll
Else:
    MsgBox "Sub cp_blk error, invalid paste modes."
End if

End Sub


'Find the end of a continuous row or column of cells
'
'End action is what happens when pressing control + arrow key, or the End key.
'It moves cell selection towards the end, but goes to infinity if the cell is
'already the end. This function does not goes to infinity, thus it is safe.
'
'Inputs:
'   cel : Cell to start.
'   dir : Direction to go to. Either: "down", "right".
'
'Returns:
'   Range of the cell where end function safely lands.
'
'Examples:
'
'Find the end of a row of cells.
'   A  B
'1  1  2
'sf_end(Range("A1"),"right")
'Range("B2")
Function sf_end( _
    ByRef cel As Range, _
    ByVal dir As String) As Range
Dim res As Range
Set res = cel
If dir = "down" Then
    If cel.Offset(1,0).Value <> "" Then
        Set res = cel.End(xlDown)
    End If
ElseIf dir = "right" Then
    If cel.Offset(0,1).Value <> "" Then 
        Set res = cel.End(xlToRight)
    End If
Else: MsgBox "Function sf_end error, invalid direction."
End If
Set sf_end = res
End Function