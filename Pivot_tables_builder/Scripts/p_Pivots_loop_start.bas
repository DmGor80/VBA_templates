Attribute VB_Name = "p_Pivots_loop_start"
Sub p_Pivots_loop()

Dim WS As Worksheet
Dim objSell As Excel.Workbook, objShab As Excel.Workbook
Dim WErr As Worksheet
Dim Final1Row As Integer
Dim Final1Col As Integer
Dim text As String
Dim cell As Range


Set objThis = Excel.ActiveWorkbook
Set WP = Worksheets("Pivots>>")
Set WS = Worksheets("DB-1-B")
oneRow = 6
name1Row = oneRow - 1
name2Row = 5 'name Row in DB sheet
name3Row = 2 'name Row in Pivots sheet
twoRow = 6
threeRow = 4 'sheet Pivots
row_col = 3


Final2Col = WS.Range("IV5").End(xlToLeft).Column
Final2Row = WS.Range("A65536").End(xlUp).Row
Final3Col = WP.Range("IV2").End(xlToLeft).Column
Final3Row = WP.Range("A65536").End(xlUp).Row

'Look a PT if it is exit
'    On Error Resume Next
'    ' Attempt to set the pivot table object
'    Set PT = WPv.PivotTables(1)
'    On Error GoTo 0
'    ' If a pivot table exists, delete it
'    If Not PT Is Nothing Then PT.TableRange2.Delete
    
' Find a Column with Columns name
j = 1
Do While WP.Cells(name3Row, j) <> "Columns"
     j = j + 1
     column_col = j
Loop
' Find a Column with Fields name
j = 5
Do While WP.Cells(name3Row, j) <> "Fields"
     j = j + 1
     field_col = j
Loop

'Loop for a PV in existing sheets
Do While threeRow <= Final3Row
row_col = 3
sheet_name = WP.Cells(threeRow, 2)
Set WPv = Worksheets(sheet_name)

'create a pivot table
Set SourseRange = WS.Cells(name2Row, 1).Resize(Final2Row - 4, Final2Col)
Set PTDestination = WPv.Cells(name1Row, 1)
'Create Name Range for the PT
WS.Names.Add Name:="PTRange", RefersTo:=SourseRange

Set PT = WPv.PivotTableWizard(SourceType:=xlDatabase, SourceData:= _
   WS.Range("PTRange"), TableDestination:=PTDestination, TableName:="Сводная таблица1")

    'Set variables for Rows
'Count Row fileds
r = row_col
     Do While Not (IsEmpty(WP.Cells(threeRow, r).Value)) And r < column_col
        r = r + 1
     Loop
    
n = 1 'number of Position in Rows PT
    Do While row_col < r
    Row1 = WP.Cells(threeRow, row_col)
    With PT.PivotFields(Row1)
        .Orientation = xlRowField
        .Position = n
    End With
    PT.PivotFields(Row1). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    n = n + 1
    row_col = row_col + 1
    Loop
    
    'variables for Columns
'Count Column Fields
r = column_col ' end of Column fields range
c = column_col ' counter of filled Coulumns in Columns Field
        Do While Not (IsEmpty(WP.Cells(threeRow, r).Value)) And r < field_col
        r = r + 1
        Loop

    n = 1 'number of Position in Rows PT
    Do While c < r
    Row1 = WP.Cells(threeRow, c)
    With PT.PivotFields(Row1)
        .Orientation = xlColumnField
        .Position = n
    End With
    n = n + 1
    c = c + 1
    Loop
' Variables for Fields
field1 = WP.Cells(threeRow, field_col)

With PT.PivotFields(field1)
    .Orientation = xlDataField
    .Function = xlSum
    .Position = 1
    .NumberFormat = "# ##0"
End With

With PT
        .HasAutoFormat = False
        .FieldListSortAscending = False
End With


Final1Col = WPv.Range("IV6").End(xlToLeft).Column
Final1Row = WPv.Range("A65536").End(xlUp).Row
PT.TableStyle2 = "PivotStyleMedium2"
WPv.Cells(name1Row, 1) = "Sum Total Value, tousends USD "
WPv.Cells(name1Row + 1, Final1Col) = "Total"
WPv.Cells(Final1Row, 1) = "Total"
Set SizeRange = WPv.Cells(1, 3).Resize(Final1Row - 2, Final1Col)
SizeRange.Select
Selection.ColumnWidth = 10.73

threeRow = threeRow + 1
Loop


End Sub


