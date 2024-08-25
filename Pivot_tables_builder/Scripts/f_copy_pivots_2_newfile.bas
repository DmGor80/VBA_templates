Attribute VB_Name = "f_copy_pivots_2_newfile"
Sub f_copy_2_newfile()
'
Dim WS As Worksheet
Dim objSell As Excel.Workbook, objShab As Excel.Workbook
Dim WErr As Worksheet
Dim Final1Row As Integer
Dim Final1Col As Integer
Dim text As String
Dim cell As Range

Set objThis = Excel.ActiveWorkbook
'Set WP = Worksheets("Pivots>>")

'Set objTest = Excel.Workbooks.Open("f:\¿Ì‡ÎËÁ\–Â‡ÎËÁ‡ˆËˇ\¿Õ¿À»«.œ–Œ√ÕŒ«\Sellers_history.xls")
Set objTest = Excel.Workbooks("Pivot_tables.xls")

oneRow = 3
name1Row = oneRow - 1
'Final1Row = WP.Range("A65536").End(xlUp).Row

name1 = ActiveSheet.Name
Set WPv = Worksheets(name1)
For Each PT In WPv.PivotTables
        ' find a first row
        PT.TableRange2.Offset(1, 0).Copy
        objTest.Activate
        Set WS = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        WS.Name = name1
        'objTest.Sheets(WS).Select
        WS.[A3].PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        ' Format this table
        Final2Row = WS.Range("A65536").End(xlUp).Row
        First2Row = WS.Range("A1").End(xlDown).Row
        Final2Col = WS.Range("IV5").End(xlToLeft).Column
        
        WS.Cells(First2Row, 1) = WS.Cells(First2Row + 1, 1)
        WS.Cells(First2Row + 1, 1).Clear
        WS.Cells(First2Row, Final2Col) = "Total"
        WS.Cells(First2Row + 1, Final2Col).Clear
        WS.Cells(Final2Row, 1) = "Total"
        
        col = Final2Col - 1
        Do While (IsEmpty(WS.Cells(First2Row, col).Value))
                col = col - 1
        Loop
        Set MRange = WS.Cells(First2Row, col).Resize(1, Final2Col - col)
        MRange.Select
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        Selection.Merge
        'put the name for the sheet
        WS.Cells(First2Row - 1, 1) = WS.Cells(First2Row, 1) & " | " & WS.Cells(First2Row + 1, 2) & " | " & WS.Cells(First2Row, col) & " | "
        WS.Cells(First2Row - 1, Final2Col) = "Annex " & name1
        
    Next PT


End Sub
