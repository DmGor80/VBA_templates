Attribute VB_Name = "code_info"
Sub Code_info_show()

Dim objSell As Excel.Workbook, objShab As Excel.Workbook
Dim ws As Worksheet
Dim WErr As Worksheet
Dim PT As PivotTable
Dim text As String
Dim Final1Row As Integer
Dim Final1Col As Integer
Dim NowDate As Long
Dim DateDiff As Integer

Set objThis = Excel.ActiveWorkbook
Set WC = Worksheets("Code_info")
Set WCdb = Worksheets("All_editions")
Set WCdbIM = Worksheets("All_editions_import") 'Must be filled manualy before (if needed)
Set WM = Worksheets("Main")

oneRow = 3
twoRow = 2
threeRow = 2
fourRow = 4

name1Row = oneRow - 1
name2Row = twoRow - 1
name3Row = threeRow - 1
name4Row = fourRow - 1

Final1Col = WC.Range("IV2").End(xlToLeft).Column
Final1Row = WC.Range("B65536").End(xlUp).Row
Final2Row = WCdb.Range("A65536").End(xlUp).Row
Final2Col = WCdb.Range("IV1").End(xlToLeft).Column
Final3Row = WCdbIM.Range("A65536").End(xlUp).Row
Final3Col = WCdbIM.Range("IV1").End(xlToLeft).Column
Final4Row = WM.Range("A65536").End(xlUp).Row
Final4Col = WM.Range("IV2").End(xlToLeft).Column

'Find columns in All_editions
c_col = WorksheetFunction.Match("CN", WCdb.Rows(name2Row), 0)
d_col = WorksheetFunction.Match("Date_of_publication", WCdb.Rows(name2Row), 0)
import_col = WorksheetFunction.Match("Import/Export", WCdb.Rows(name2Row), 0)
annex_col = WorksheetFunction.Match("Annex", WCdb.Rows(name2Row), 0)
article_col = WorksheetFunction.Match("Article", WCdb.Rows(name2Row), 0)

'Find columns in Codes
next_digit_col1 = WorksheetFunction.Match("Code", WC.Rows(name1Row), 0)
date_col = WorksheetFunction.Match("Date of Publication", WC.Rows(name1Row), 0)
annex_column = WorksheetFunction.Match("Annex", WC.Rows(name1Row), 0)
article_column = annex_column + 1
import_column = WorksheetFunction.Match("Import to RU/Export from RU", WC.Rows(name1Row), 0)
status_col = WorksheetFunction.Match("Status", WC.Rows(name1Row), 0)
rownb_col = WorksheetFunction.Match("Row number in All_editions sheet", WC.Rows(name1Row), 0)

'Find columns in Main
code_col = WorksheetFunction.Match("HS Code", WM.Rows(name4Row), 0)
high_priority_col = WorksheetFunction.Match("High-priority Items (last edition) (Yes, No)", WM.Rows(name4Row), 0)
weapon_col = WorksheetFunction.Match("Weapon  (Yes, No)", WM.Rows(name4Row), 0)
transit_col = WorksheetFunction.Match("Transit prohibited (last edition) (Yes, No)", WM.Rows(name4Row), 0)

Set HSRangeWCdb = WCdb.Cells(twoRow, c_col).Resize(Final2Row - 2, d_col)
RowCount = HSRangeWCdb.Rows.Count
ColCount = HSRangeWCdb.Columns.Count

'Clear the previous content
If oneRow <= Final1Row Then
Set ClearRange = WC.Range(WC.Cells(oneRow, 2), WC.Cells(Final1Row, Final1Col))
ClearRange.ClearContents
With ClearRange.Font
        '.Color = xlAutomatic
        .TintAndShade = 0 ' Reset tint and shade to 0
        .Bold = False ' Reset bold property to False
End With
ClearRange.Interior.Color = xlNone
End If


'GoTo Test_Last
'Check the formulas in XXX columns
hscode = Val(WC.Cells(3, 1))
'XXXX-XXXX
form = Val(Left(hscode, 8))
startRow = twoRow
Do While startRow < Final2Row
    Set xLookupArray = WCdb.Range(WCdb.Cells(startRow, 1), WCdb.Cells(Final2Row, 1))
    On Error Resume Next
    rowNum = Application.Match(form, xLookupArray, 0)
    On Error GoTo 0 ' Reset error handling
    ' Check if MATCH returned an error (no match found)
    If IsError(rowNum) Then
        rowNum = Final2Row + 1
        Exit Do
    End If
    rowNum = rowNum + startRow - 1
    Formula = "=IFERROR(INDEX(All_editions!$A$" & startRow & ":$A$" & Final2Row & ",MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)),"""")"
    Formula23 = "=MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)+" & startRow - 1
    WC.Cells(oneRow, next_digit_col1).Formula = Formula
    WC.Cells(oneRow, rownb_col).Formula = Formula23
    adress = "'" & WCdb.Name & "'" & "!R" & rowNum & "C" & c_col
    text = WC.Cells(oneRow, rownb_col)
    WC.Cells(oneRow, rownb_col).Hyperlinks.Add Anchor:=WC.Cells(oneRow, rownb_col), Address:="", SubAddress:= _
        adress, TextToDisplay:=text
    If WC.Cells(oneRow, next_digit_col1) <> "" Then
        WC.Cells(oneRow, annex_column) = WCdb.Cells(rowNum, annex_col)
        WC.Cells(oneRow, article_column) = WCdb.Cells(rowNum, article_col)
        WC.Cells(oneRow, import_column) = WCdb.Cells(rowNum, import_col)
        WC.Cells(oneRow, date_col) = WCdb.Cells(rowNum, d_col)
        WC.Cells(oneRow, status_col) = "2-Likely banned"
        startRow = rowNum + 2
        oneRow = oneRow + 1
    End If
Loop

'XXXX-XX-00
form_previous = form
form = Val(Left(hscode, 6) & "00")
startRow = twoRow
Do While startRow < Final2Row
    
    If form = form_previous Then Exit Do 'Check if XXXX-XXXX=XXXX-XX-00 to avoid doublerows
    
    Set xLookupArray = WCdb.Range(WCdb.Cells(startRow, 1), WCdb.Cells(Final2Row, 1))
    On Error Resume Next
    rowNum = Application.Match(form, xLookupArray, 0)
    On Error GoTo 0 ' Reset error handling
    ' Check if MATCH returned an error (no match found)
    If IsError(rowNum) Then
        rowNum = Final2Row + 1
        Exit Do
    End If
    rowNum = rowNum + startRow - 1
    Formula = "=IFERROR(INDEX(All_editions!$A$" & startRow & ":$A$" & Final2Row & ",MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)),"""")"
    Formula23 = "=MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)+" & startRow - 1
    WC.Cells(oneRow, next_digit_col1).Formula = Formula
    WC.Cells(oneRow, rownb_col).Formula = Formula23
    adress = "'" & WCdb.Name & "'" & "!R" & rowNum & "C" & c_col
    text = WC.Cells(oneRow, rownb_col)
    WC.Cells(oneRow, rownb_col).Hyperlinks.Add Anchor:=WC.Cells(oneRow, rownb_col), Address:="", SubAddress:= _
        adress, TextToDisplay:=text
    If WC.Cells(oneRow, next_digit_col1) <> "" Then
        WC.Cells(oneRow, annex_column) = WCdb.Cells(rowNum, annex_col)
        WC.Cells(oneRow, article_column) = WCdb.Cells(rowNum, article_col)
        WC.Cells(oneRow, import_column) = WCdb.Cells(rowNum, import_col)
        WC.Cells(oneRow, date_col) = WCdb.Cells(rowNum, d_col)
        WC.Cells(oneRow, status_col) = "2-Likely banned"
        startRow = rowNum + 2
        oneRow = oneRow + 1
    End If
Loop
'XXXX-0000
form_previous = form
form = Val(Left(hscode, 4) & "0000")
startRow = twoRow
Do While startRow < Final2Row

    If form = form_previous Then Exit Do 'Check if XXXX-XX-00=XXXX-0000 to avoid doublerows
    
    Set xLookupArray = WCdb.Range(WCdb.Cells(startRow, 1), WCdb.Cells(Final2Row, 1))
    On Error Resume Next
    rowNum = Application.Match(form, xLookupArray, 0)
    On Error GoTo 0 ' Reset error handling
    ' Check if MATCH returned an error (no match found)
    If IsError(rowNum) Then
        rowNum = Final2Row + 1
        Exit Do
    End If
    rowNum = rowNum + startRow - 1
    Formula = "=IFERROR(INDEX(All_editions!$A$" & startRow & ":$A$" & Final2Row & ",MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)),"""")"
    Formula23 = "=MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)+" & startRow - 1
    WC.Cells(oneRow, next_digit_col1).Formula = Formula
    WC.Cells(oneRow, rownb_col).Formula = Formula23
    adress = "'" & WCdb.Name & "'" & "!R" & rowNum & "C" & c_col
    text = WC.Cells(oneRow, rownb_col)
    WC.Cells(oneRow, rownb_col).Hyperlinks.Add Anchor:=WC.Cells(oneRow, rownb_col), Address:="", SubAddress:= _
        adress, TextToDisplay:=text
    If WC.Cells(oneRow, next_digit_col1) <> "" Then
        WC.Cells(oneRow, annex_column) = WCdb.Cells(rowNum, annex_col)
        WC.Cells(oneRow, article_column) = WCdb.Cells(rowNum, article_col)
        WC.Cells(oneRow, import_column) = WCdb.Cells(rowNum, import_col)
        WC.Cells(oneRow, date_col) = WCdb.Cells(rowNum, d_col)
        WC.Cells(oneRow, status_col) = "2-Likely banned"
        startRow = rowNum + 2
        oneRow = oneRow + 1
    End If
Loop

'XX
form = Val(Left(hscode, 2))
startRow = twoRow
Do While startRow < Final2Row
    Set xLookupArray = WCdb.Range(WCdb.Cells(startRow, 1), WCdb.Cells(Final2Row, 1))
    On Error Resume Next
    rowNum = Application.Match(form, xLookupArray, 0)
    On Error GoTo 0 ' Reset error handling
    ' Check if MATCH returned an error (no match found)
    If IsError(rowNum) Then
        rowNum = Final2Row + 1
        Exit Do
    End If
    rowNum = rowNum + startRow - 1
    Formula = "=IFERROR(INDEX(All_editions!$A$" & startRow & ":$A$" & Final2Row & ",MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)),"""")"
    Formula23 = "=MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)+" & startRow - 1
    WC.Cells(oneRow, next_digit_col1).Formula = Formula
    WC.Cells(oneRow, rownb_col).Formula = Formula23
    adress = "'" & WCdb.Name & "'" & "!R" & rowNum & "C" & c_col
    text = WC.Cells(oneRow, rownb_col)
    WC.Cells(oneRow, rownb_col).Hyperlinks.Add Anchor:=WC.Cells(oneRow, rownb_col), Address:="", SubAddress:= _
        adress, TextToDisplay:=text
    If WC.Cells(oneRow, next_digit_col1) <> "" Then
        WC.Cells(oneRow, annex_column) = WCdb.Cells(rowNum, annex_col)
        WC.Cells(oneRow, article_column) = WCdb.Cells(rowNum, article_col)
        WC.Cells(oneRow, import_column) = WCdb.Cells(rowNum, import_col)
        WC.Cells(oneRow, date_col) = WCdb.Cells(rowNum, d_col)
        WC.Cells(oneRow, status_col) = "1-Banned"
        startRow = rowNum + 2
        oneRow = oneRow + 1
    End If
Loop

'XXX
form = Val(Left(hscode, 3))
startRow = twoRow
Do While startRow < Final2Row
    Set xLookupArray = WCdb.Range(WCdb.Cells(startRow, 1), WCdb.Cells(Final2Row, 1))
    On Error Resume Next
    rowNum = Application.Match(form, xLookupArray, 0)
    On Error GoTo 0 ' Reset error handling
    ' Check if MATCH returned an error (no match found)
    If IsError(rowNum) Then
        rowNum = Final2Row + 1
        Exit Do
    End If
    rowNum = rowNum + startRow - 1
    Formula = "=IFERROR(INDEX(All_editions!$A$" & startRow & ":$A$" & Final2Row & ",MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)),"""")"
    Formula23 = "=MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)+" & startRow - 1
    WC.Cells(oneRow, next_digit_col1).Formula = Formula
    WC.Cells(oneRow, rownb_col).Formula = Formula23
    adress = "'" & WCdb.Name & "'" & "!R" & rowNum & "C" & c_col
    text = WC.Cells(oneRow, rownb_col)
    WC.Cells(oneRow, rownb_col).Hyperlinks.Add Anchor:=WC.Cells(oneRow, rownb_col), Address:="", SubAddress:= _
        adress, TextToDisplay:=text
    If WC.Cells(oneRow, next_digit_col1) <> "" Then
        WC.Cells(oneRow, annex_column) = WCdb.Cells(rowNum, annex_col)
        WC.Cells(oneRow, article_column) = WCdb.Cells(rowNum, article_col)
        WC.Cells(oneRow, import_column) = WCdb.Cells(rowNum, import_col)
        WC.Cells(oneRow, date_col) = WCdb.Cells(rowNum, d_col)
        WC.Cells(oneRow, status_col) = "1-Banned"
        startRow = rowNum + 2
        oneRow = oneRow + 1
    End If
Loop

'XXXX
form = Val(Left(hscode, 4))
startRow = twoRow
Do While startRow < Final2Row
    Set xLookupArray = WCdb.Range(WCdb.Cells(startRow, 1), WCdb.Cells(Final2Row, 1))
    On Error Resume Next
    rowNum = Application.Match(form, xLookupArray, 0)
    On Error GoTo 0 ' Reset error handling
    ' Check if MATCH returned an error (no match found)
    If IsError(rowNum) Then
        rowNum = Final2Row + 1
        Exit Do
    End If
    rowNum = rowNum + startRow - 1
    Formula = "=IFERROR(INDEX(All_editions!$A$" & startRow & ":$A$" & Final2Row & ",MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)),"""")"
    Formula23 = "=MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)+" & startRow - 1
    WC.Cells(oneRow, next_digit_col1).Formula = Formula
    WC.Cells(oneRow, rownb_col).Formula = Formula23
    adress = "'" & WCdb.Name & "'" & "!R" & rowNum & "C" & c_col
    text = WC.Cells(oneRow, rownb_col)
    WC.Cells(oneRow, rownb_col).Hyperlinks.Add Anchor:=WC.Cells(oneRow, rownb_col), Address:="", SubAddress:= _
        adress, TextToDisplay:=text
    If WC.Cells(oneRow, next_digit_col1) <> "" Then
        WC.Cells(oneRow, annex_column) = WCdb.Cells(rowNum, annex_col)
        WC.Cells(oneRow, article_column) = WCdb.Cells(rowNum, article_col)
        WC.Cells(oneRow, import_column) = WCdb.Cells(rowNum, import_col)
        WC.Cells(oneRow, date_col) = WCdb.Cells(rowNum, d_col)
        WC.Cells(oneRow, status_col) = "1-Banned"
        startRow = rowNum + 2
        oneRow = oneRow + 1
    End If
Loop

'XXXX-X
form = Val(Left(hscode, 5))
startRow = twoRow
Do While startRow < Final2Row
    Set xLookupArray = WCdb.Range(WCdb.Cells(startRow, 1), WCdb.Cells(Final2Row, 1))
    On Error Resume Next
    rowNum = Application.Match(form, xLookupArray, 0)
    On Error GoTo 0 ' Reset error handling
    ' Check if MATCH returned an error (no match found)
    If IsError(rowNum) Then
        rowNum = Final2Row + 1
        Exit Do
    End If
    rowNum = rowNum + startRow - 1
    Formula = "=IFERROR(INDEX(All_editions!$A$" & startRow & ":$A$" & Final2Row & ",MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)),"""")"
    Formula23 = "=MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)+" & startRow - 1
    WC.Cells(oneRow, next_digit_col1).Formula = Formula
    WC.Cells(oneRow, rownb_col).Formula = Formula23
    adress = "'" & WCdb.Name & "'" & "!R" & rowNum & "C" & c_col
    text = WC.Cells(oneRow, rownb_col)
    WC.Cells(oneRow, rownb_col).Hyperlinks.Add Anchor:=WC.Cells(oneRow, rownb_col), Address:="", SubAddress:= _
        adress, TextToDisplay:=text
    If WC.Cells(oneRow, next_digit_col1) <> "" Then
        WC.Cells(oneRow, annex_column) = WCdb.Cells(rowNum, annex_col)
        WC.Cells(oneRow, article_column) = WCdb.Cells(rowNum, article_col)
        WC.Cells(oneRow, import_column) = WCdb.Cells(rowNum, import_col)
        WC.Cells(oneRow, date_col) = WCdb.Cells(rowNum, d_col)
        WC.Cells(oneRow, status_col) = "1-Banned"
        startRow = rowNum + 2
        oneRow = oneRow + 1
    End If
Loop

'XXXX-XX
form = Val(Left(hscode, 6))
startRow = twoRow
Do While startRow < Final2Row
    Set xLookupArray = WCdb.Range(WCdb.Cells(startRow, 1), WCdb.Cells(Final2Row, 1))
    On Error Resume Next
    rowNum = Application.Match(form, xLookupArray, 0)
    On Error GoTo 0 ' Reset error handling
    ' Check if MATCH returned an error (no match found)
    If IsError(rowNum) Then
        rowNum = Final2Row + 1
        Exit Do
    End If
    rowNum = rowNum + startRow - 1
    Formula = "=IFERROR(INDEX(All_editions!$A$" & startRow & ":$A$" & Final2Row & ",MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)),"""")"
    Formula23 = "=MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)+" & startRow - 1
    WC.Cells(oneRow, next_digit_col1).Formula = Formula
    WC.Cells(oneRow, rownb_col).Formula = Formula23
    adress = "'" & WCdb.Name & "'" & "!R" & rowNum & "C" & c_col
    text = WC.Cells(oneRow, rownb_col)
    WC.Cells(oneRow, rownb_col).Hyperlinks.Add Anchor:=WC.Cells(oneRow, rownb_col), Address:="", SubAddress:= _
        adress, TextToDisplay:=text
    If WC.Cells(oneRow, next_digit_col1) <> "" Then
        WC.Cells(oneRow, annex_column) = WCdb.Cells(rowNum, annex_col)
        WC.Cells(oneRow, article_column) = WCdb.Cells(rowNum, article_col)
        WC.Cells(oneRow, import_column) = WCdb.Cells(rowNum, import_col)
        WC.Cells(oneRow, date_col) = WCdb.Cells(rowNum, d_col)
        WC.Cells(oneRow, status_col) = "1-Banned"
        startRow = rowNum + 2
        oneRow = oneRow + 1
    End If
Loop
'XXXX-XXX
form = Val(Left(hscode, 7))
startRow = twoRow
Do While startRow < Final2Row
    Set xLookupArray = WCdb.Range(WCdb.Cells(startRow, 1), WCdb.Cells(Final2Row, 1))
    On Error Resume Next
    rowNum = Application.Match(form, xLookupArray, 0)
    On Error GoTo 0 ' Reset error handling
    ' Check if MATCH returned an error (no match found)
    If IsError(rowNum) Then
        rowNum = Final2Row + 1
        Exit Do
    End If
    rowNum = rowNum + startRow - 1
    Formula = "=IFERROR(INDEX(All_editions!$A$" & startRow & ":$A$" & Final2Row & ",MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)),"""")"
    Formula23 = "=MATCH(" & form & ",All_editions!$A$" & startRow & ":$A$" & Final2Row & ",0)+" & startRow - 1
    WC.Cells(oneRow, next_digit_col1).Formula = Formula
    WC.Cells(oneRow, rownb_col).Formula = Formula23
    adress = "'" & WCdb.Name & "'" & "!R" & rowNum & "C" & c_col
    text = WC.Cells(oneRow, rownb_col)
    WC.Cells(oneRow, rownb_col).Hyperlinks.Add Anchor:=WC.Cells(oneRow, rownb_col), Address:="", SubAddress:= _
        adress, TextToDisplay:=text
    If WC.Cells(oneRow, next_digit_col1) <> "" Then
        WC.Cells(oneRow, annex_column) = WCdb.Cells(rowNum, annex_col)
        WC.Cells(oneRow, article_column) = WCdb.Cells(rowNum, article_col)
        WC.Cells(oneRow, import_column) = WCdb.Cells(rowNum, import_col)
        WC.Cells(oneRow, date_col) = WCdb.Cells(rowNum, d_col)
        WC.Cells(oneRow, status_col) = "3-Undefined"
        startRow = rowNum + 2
        oneRow = oneRow + 1
    End If
Loop

'Check if there is some Data or the code was "4-Not banned"
oneRow = name1Row + 1
If IsEmpty(WC.Cells(oneRow, status_col).value) Then
WC.Cells(oneRow, status_col) = "This Code doesn't exist in DataBase"
Exit Sub
End If

'Design Block
'Sorting

Final1Col = WC.Range("IV2").End(xlToLeft).Column
Final1Row = WC.Range("B65536").End(xlUp).Row
Set SortRange = WC.Range(WC.Cells(oneRow, 2), WC.Cells(Final1Row, Final1Col))
Set ImportRanges = WC.Range(WC.Cells(oneRow, import_column), WC.Cells(Final1Row, import_column))
Set StatusRange = WC.Range(WC.Cells(oneRow, status_col), WC.Cells(Final1Row, status_col))
Set DateRange = WC.Range(WC.Cells(oneRow, date_col), WC.Cells(Final1Row, date_col))
' Sort the SortRange
'SortRange.Select
With SortRange
    .Sort Key1:=ImportRanges, Order1:=xlDescending, _
          Key2:=StatusRange, Order2:=xlDescending, _
          Key3:=DateRange, Order3:=xlAscending, _
          Header:=xlNo, MatchCase:=False, Orientation:=xlTopToBottom, SortMethod:=xlPinYin
End With
SortRange.HorizontalAlignment = xlCenter
SortRange.VerticalAlignment = xlCenter
'Coloring
importRow = name1Row
exportRow = importRow
otherRow = exportRow
For i = otherRow To Final1Row
    If WC.Cells(i, import_column) = "Other" Then
    otherRow = i
    End If
Next i
For i = importRow To Final1Row
    If WC.Cells(i, import_column) = "Import" Then
    importRow = i
    End If
Next i
For i = importRow To Final1Row
    If WC.Cells(i, import_column) = "Export" Then
    exportRow = i
    End If
Next i

If otherRow > 2 Then
Set OtherRange = WC.Range(WC.Cells(oneRow, 2), WC.Cells(otherRow, Final1Col))
With OtherRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
End With
End If
If importRow > 2 Then
    If otherRow = 2 Then
    Set ImportRange = WC.Range(WC.Cells(oneRow, 2), WC.Cells(importRow, Final1Col))
    Else
    Set ImportRange = WC.Range(WC.Cells(otherRow + 1, 2), WC.Cells(importRow, Final1Col))
    End If
With ImportRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.899990844447157
        .PatternTintAndShade = 0
End With
End If


If exportRow > 2 Then
    If importRow = 2 Then
    Set ExportRange = WC.Range(WC.Cells(oneRow, 2), WC.Cells(exportRow, Final1Col))
    Else
    Set ExportRange = WC.Range(WC.Cells(importRow + 1, 2), WC.Cells(exportRow, Final1Col))
    End If
With ExportRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
End With
End If
'Balding
If IsObject(ImportRange) Then
    If Not ImportRange Is Nothing Then
        ' Format the first row of ImportRange as bold
        ImportRange.Rows(1).Font.Bold = True
        Do While oneRow + 1 <= importRow
            If WC.Cells(oneRow, status_col) <> WC.Cells(oneRow + 1, status_col) Then
            WC.Range(WC.Cells(oneRow + 1, 2), WC.Cells(oneRow + 1, Final1Col)).Font.Bold = True
            Exit Do
            End If
        oneRow = oneRow + 1
        Loop
   End If
End If


End Sub
