Attribute VB_Name = "p_PivotStyleUpdate"
Sub p_PivotStyleUpdate()
Attribute p_PivotStyleUpdate.VB_ProcData.VB_Invoke_Func = " \n14"
'
Dim WS As Worksheet
Dim objSell As Excel.Workbook, objShab As Excel.Workbook
Dim WErr As Worksheet
Dim Final1Row As Integer
Dim Final1Col As Integer
Dim text As String
Dim cell As Range



Set objThis = Excel.ActiveWorkbook
Set WP = Worksheets("Pivots>>")

oneRow = 3
name1Row = oneRow - 1
Final1Row = WP.Range("A65536").End(xlUp).Row

'
Do While oneRow <= Final1Row
name1 = WP.Cells(oneRow, 2)
Set WPv = Worksheets(name1)
For Each PT In WPv.PivotTables
        ' Change the PivotTable style
        PT.TableStyle2 = "PivotStyleLight20"
    Next PT
'
oneRow = oneRow + 1
Loop


End Sub
