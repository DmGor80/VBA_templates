Attribute VB_Name = "code_info_call"
Public hscodes As String

Sub Code_info_CALL()

Dim objSell As Excel.Workbook, objShab As Excel.Workbook
Dim ws As Worksheet
Dim WErr As Worksheet
Dim PT As PivotTable
Dim text As String
Dim Final1Row As Integer
Dim Final1Col As Integer
Dim NowDate As Long
Dim DateDiff As Integer
'Dim hscodes As String

Set objThis = Excel.ActiveWorkbook
Set WC = Worksheets("Code_info")
Set WM = Worksheets("Main")

oneRow = 3
fourRow = 4
name1Row = oneRow - 1
name4Row = fourRow - 1

Final1Col = WC.Range("IV2").End(xlToLeft).Column
Final1Row = WC.Range("B65536").End(xlUp).Row
Final4Row = WM.Range("A65536").End(xlUp).Row
Final4Col = WM.Range("IV2").End(xlToLeft).Column

'Find columns in Main
code_col = WorksheetFunction.Match("HS Code", WM.Rows(name4Row), 0)

ActiveSheet.Activate
aRow = ActiveCell.Row
hscodes = ActiveSheet.Cells(aRow, code_col)
If IsEmpty(hscodes) Or Len(hscodes) <> 10 Then
  frmWrongHsCode.Show
  'Exit Sub
End If

frmHsCodeCorrect.HSCodelbl.Caption = hscodes

frmHsCodeCorrect.Show

WC.Cells(oneRow, 1) = hscodes

Call Code_info_show

WC.Activate

End Sub

