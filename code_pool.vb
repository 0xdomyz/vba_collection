Application.StatusBar = "asdf"

with rng
    .AutoFilter Field:=1, Criteria1:=">=" & varname, Operator:=xlAnd, _
    Criteria2:="<=" & varname

ReDim varname(0)
varname(0) = ""

On Error Resume Next
Do While rng <> ""
Loop

ReDim Preserve

On Error GoTo 0

rng.Sort Key1:=rng, Order1:=xlAscending, Header:=xlYes, OrderCustom:=1, _
MatchCase:=True, Orientation:=xlTopToBottom, DataOption1:=xlSortTextAsNumbers

IsEmpty(rng.Value)

rngs As Variant
For Each itm In rngs
Next itm

Exit Sub

rng.Select
Selection.AutoFill Destination:=rng

ThisWorkbook.Worksheets("sht").Cells.EntireColum.AutoFit

Function col2arr(ByRef rng As Range) As Variant
    Dim i As Integer
    Dim arr() As Variant
    i = 0
    Do While rng.Value <> ""
        ReDim Preserve arr(i)
        arr(i) = rng.Value
        i = i + 1
        Set rng = rng.Offset(1,0)
    Loop
    col2arr = arr
End Function


