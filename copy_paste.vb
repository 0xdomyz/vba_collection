'Assign values from range a to range b
'
've_blk(a,b,pst_mod):
'Range a represented by top left corner cell of a continuous block.
'Continuous means no empty cells on top and left edges.
'Range b has same dimension as range a.
'
've_wh(a,w,h,b,pst_mod):
'Ranges defined by top left corner cells and their width and height.
'
've(a,b,pst_mod):
'Specify ranges fully.
'
'Inputs:
'   a : Top left cell of the original block. Original range for ve.
'   w : Width of the table.
'   h : Height of the table.
'   b : Top left cell of the target location. Target range for ve.
'
'Example
'
'Assign value for a block of cells from A1 to E1
'   A  B  C
'1  1  2
'2  3  4
'3     5  6
'Call ve_blk(Range("A1"), Range("E1"))
'Call ve_wh(Range("A1"), 2, 2, Range("E1"))
'Call ve(Range("A1:B2"), Range("E1"))
'   E  F  G
'1  1  2
'2  3  4
'3
Sub ve_blk(ByRef a As Range, ByRef b As Range)
    Dim w As Integer
    Dim h As Integer
    w = sf_end(a, "right").Column - a.Column + 1
    h = sf_end(a, "down").Row - a.Row + 1
    Call ve_wh(a,w,h,b)
End Sub
Sub ve_wh( _
        ByRef a As Range, _
        ByVal w As Integer, _
        ByVal h As Integer, _
        ByRef b As Range)
    Call ve(Range(a, a.Offset(h - 1, w - 1)),Range(b, b.Offset(h - 1, w - 1)))
End Sub
Sub ve(ByRef a As Range, ByRef b As Range)
    b.value = a.value
End Sub

'Copy paste range a to range b
'
'cp_blk(a,b,pst_mod):
'Range a represented by top left corner cell of a continuous block.
'Continuous means no empty cells on top and left edges.
'
'cp_wh(a,w,h,b,pst_mod):
'Range a defined by top left corner cell and it's width and height.
'
'cp(a,b,pst_mod):
'Specify range a fully.
'
'Inputs:
'   a : Top left cell of the original block. Original range for cp.
'   w : Width of the table.
'   h : Height of the table.
'   b : Top left corner of the target location.
'   pst_mod : Optional Excel paste mode integer.
'       Default : xlPasteValues
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
'Call cp_blk(Range("A1"), Range("E1"))
'Call cp_wh(Range("A1"), 2, 2, Range("E1"))
'Call cp(Range("A1:B2"), Range("E1"))
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
Sub cp_wh( _
        ByRef a As Range, _
        ByVal w As Integer, _
        ByVal h As Integer, _
        ByRef b As Range, _
        Optional ByVal pst_mod As Integer = xlPasteValues)
    Call cp(Range(a, a.Offset(h - 1, w - 1)),b,pst_mod)
End Sub
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