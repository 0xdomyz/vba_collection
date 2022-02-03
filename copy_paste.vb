'Copy paste a block of cells
'
'Copy a continuous row, column, or table of cells.
'Continuous means no empty cells on top and left edges of the block.
'Then paste into another location represented by a cell. Top left corner of the 
'block is to be placed into this representative cell.
'
'Inputs:
'   a : Top left cell of the original block.
'   b : Representative cell of the target location.
'   pst_mod : Optional Excel paste mode integer. See sub cp for details.
'       Default : xlPasteValues
'
'Effects:
'   A copy paste action.
'
'Example
'
'Copy paste a block of cells from A1 to E1
'   A  B  C
'1  1  2
'2  3  4
'3     5  6
'Call cp_blk(Range("A1"), Range("E1"))
'   E  F  G
'1  1  2
'2  3  4
'3
Sub cp_blk( _
    ByRef a As Range, _
    ByRef b As Range, _
    Optional ByVal pst_mod As Integer = xlPasteValues)
Dim w As Integer
Dim h As Integer
w = sf_end(a, "right").Column - a.Column + 1
h = sf_end(a, "down").Row - a.Row + 1
Call cp_wh(a,w,h,b,pst_mod)
End Sub

'Copy paste a range defined by dimension
'
'Dimension is the width and height of the table being copied.
'Then paste into another location represented by a cell. Top left corner of the 
'original table is to be placed into this representative cell.
'
'Inputs:
'   a : Top left cell of the original table.
'   w : Width of the table.
'   h : Height of the table.
'   b : Representative cell of the target location.
'   pst_mod : Optional Excel paste mode integer. See sub cp for details.
'       Default : xlPasteValues
'
'Effects:
'   A copy paste action.
'
'Example
'
'Copy paste a block of cells from A1 to E1
'   A  B  C
'1  1  2
'2  3  4
'3     5  6
'Call cp_wh(Range("A1"), 2, 2, Range("E1"))
'   E  F  G
'1  1  2
'2  3  4
'3
Sub cp_wh( _
    ByRef a As Range, _
    ByRef w As Integer, _
    ByRef h As Integer, _
    ByRef b As Range, _
    Optional ByVal pst_mod As Integer = xlPasteValues)
Dim br As Range
Set br = a.Offset(h - 1, w - 1)
Call cp(Range(a, br),b,pst_mod)
End Sub

'Copy paste a range
'
'Copy a range.
'Then paste into another location represented by a cell. Top left corner of the 
'original range is to be placed into this representative cell.
'
'Inputs:
'   a : Original range.
'   b : Representative cell of the target location.
'   pst_mod : Optional Excel paste mode integer.
'       Default : xlPasteValues
'
'Effects:
'   A copy paste action.
'
'Notes:
'List of Excel paste modes:
'   https://docs.microsoft.com/en-us/office/vba/api/Excel.XlPasteType
'Commonly used:
'   xlPasteValues, xlPasteValuesAndNumberFormats
'   ,xlPasteFormats, xlPasteColumnWidths
'   ,xlPasteAll
'
'Example
'
'Copy paste a block of cells from A1 to E1
'   A  B  C
'1  1  2
'2  3  4
'3     5  6
'Call cp(Range("A1:B2"), Range("E1"))
'   E  F  G
'1  1  2
'2  3  4
'3
Sub cp( _
    ByRef a As Range, _
    ByRef b As Range, _
    Optional ByVal pst_mod As Integer = xlPasteValues)
a.Copy
b.PasteSpecial Paste:=pst_mod, _
Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End Sub

'Find range from a safely done end command
'
'End command is what happens when pressing control + arrow key, or the End key.
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