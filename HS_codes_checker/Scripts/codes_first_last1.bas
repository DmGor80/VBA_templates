Attribute VB_Name = "codes_first_last1"
Sub Codes_first_last()

Dim objSell As Excel.Workbook
Dim WC As Worksheet
Dim WE As Worksheet
Dim WM As Worksheet
Dim WCdb As Worksheet

Dim Final1Row As Integer
Dim Final1Col As Integer

Dim originalRange As Range
Dim cell As Range
Dim uniqueValues As Object
Dim newRange As Range
Dim value As Variant

Set objThis = Excel.ActiveWorkbook
Set WC = Worksheets("Codes_first_last")
Set WCdb = Worksheets("All_editions")
Set WE = Worksheets("Editions")
Set WM = Worksheets("Main")
Set WCdbIM = Worksheets("All_editions_import") 'Must be filled manualy before (if needed)
Set WLIM = Worksheets("Last_edition_import") 'Changes automaticly
oneRow = 3
twoRow = 2
threeRow = 2
fourRow = 4
fiveRow = 2
sixRow = 2
name1Row = oneRow - 1
name2Row = twoRow - 1
name3Row = threeRow - 1
name4Row = fourRow - 1
name5Row = fiveRow - 1
name6Row = sixRow - 1

Final1Col = WC.Range("IV2").End(xlToLeft).Column
Final1Row = WC.Range("B65536").End(xlUp).Row
Final2Row = WCdb.Range("A65536").End(xlUp).Row
Final2Col = WCdb.Range("IV1").End(xlToLeft).Column
Final3Row = WE.Range("A65536").End(xlUp).Row
Final3Col = WE.Range("IV1").End(xlToLeft).Column
Final4Row = WM.Range("A65536").End(xlUp).Row
Final4Col = WM.Range("IV2").End(xlToLeft).Column
Final5Row = WCdbIM.Range("A65536").End(xlUp).Row
Final5Col = WCdbIM.Range("IV1").End(xlToLeft).Column
Final6Row = WLIM.Range("B65536").End(xlUp).Row
Final6Col = WLIM.Range("IV1").End(xlToLeft).Column


'Find columns in All_editions
c_col = WorksheetFunction.Match("CN", WCdb.Rows(name2Row), 0)
d_col = WorksheetFunction.Match("Date_of_publication", WCdb.Rows(name2Row), 0)
import_col = WorksheetFunction.Match("Import/Export", WCdb.Rows(name2Row), 0)

'Find columns in Codes
first_col = WorksheetFunction.Match("First Editions Result", WC.Rows(name1Row), 0)
ban_col1 = first_col - 2
first_date_col = first_col + 1
last_col = WorksheetFunction.Match("Last Editions Result", WC.Rows(name1Row), 0)
ban_col2 = last_col - 1
last_date_col = last_col + 1
last_annex = WorksheetFunction.Match("Last Edition Annex", WC.Rows(name1Row), 0)
last_article = last_annex + 1
'Find columns in Editions
date_edition_col = WorksheetFunction.Match("Edition's date", WE.Rows(name3Row), 0)
'Find columns in Main
code_col = WorksheetFunction.Match("HS Code", WM.Rows(name4Row), 0)
date_trans_col = WorksheetFunction.Match("Date", WM.Rows(name4Row), 0)
f_date_col = WorksheetFunction.Match("First Editions Result", WM.Rows(name4Row), 0)
t_date_col = WorksheetFunction.Match("Transaction's date Result (Grace period is ignored)", WM.Rows(name4Row), 0)
l_date_col = WorksheetFunction.Match("Last Editions Result", WM.Rows(name4Row), 0)
l_annex = WorksheetFunction.Match("Last Edition Annex", WM.Rows(name4Row), 0)
l_article = l_annex + 1

' Set the password for the protected sheet
Password = "U82024" ' Replace with the actual password
' Reference the protected worksheet
WC.Unprotect Password

'GoTo Test_Last
WM.AutoFilterMode = False
'Clean the prev. results from Codes
WC.Range(WC.Cells(oneRow, 3), WC.Cells(oneRow + 3000, Final1Col)).Font.ColorIndex = xlAutomatic
WC.Range(WC.Cells(oneRow, 2), WC.Cells(oneRow + 3000, 2)).ClearContents 'Del HS Code
WC.Range(WC.Cells(oneRow, first_col), WC.Cells(oneRow + 3000, first_col)).ClearContents
WC.Range(WC.Cells(oneRow, last_col), WC.Cells(oneRow + 3000, last_col)).ClearContents
WC.Range(WC.Cells(oneRow, last_col + 2), WC.Cells(oneRow + 3000, Final1Col)).ClearContents
WM.Range(WM.Cells(fourRow, 3), WM.Cells(fourRow + 13000, Final4Col)).Font.ColorIndex = xlAutomatic
WM.Range(WM.Cells(fourRow, 3), WM.Cells(fourRow + 13000, Final4Col)).ClearContents


' Disable screen updating to prevent flickering
Application.ScreenUpdating = False

' Create a Unique HS Code
' Define your original range
Set originalRange = WM.Range(WM.Cells(fourRow, code_col), WM.Cells(Final4Row, code_col))
' Initialize a scripting dictionary to track unique values
Set uniqueValues = CreateObject("Scripting.Dictionary")
' Loop through each cell in the original range
    For Each cell In originalRange
        ' Check if the cell value is not already in the dictionary
        If Not uniqueValues.Exists(cell.value) Then
            ' Add the cell value to the dictionary
            uniqueValues.Add cell.value, Nothing
        End If
    Next cell
WC.Cells(oneRow, 2).Resize(uniqueValues.Count, 1).value = Application.Transpose(uniqueValues.keys)
Final1Col = WC.Range("IV2").End(xlToLeft).Column
Final1Row = WC.Range("B65536").End(xlUp).Row

test_first:

'Fill the sheet WLIM
' Reference the protected worksheet
WLIM.Unprotect Password

WLIM.Range(WLIM.Cells(sixRow, 1), WLIM.Cells(Final6Row, Final6Col)).Clear
threeRow = 2
WM.Activate
'take a current date from WM
now_date = Date
    
WE.Activate
Do While threeRow <= Final3Row
    If WE.Cells(threeRow, date_edition_col) < now_date Then
    edition_date = WE.Cells(threeRow, date_edition_col)
    Else
        Exit Do
    End If
threeRow = threeRow + 1
Loop

   'Find the codes
twoRow = 2
oneRow = name1Row + 1

' Find the diapasone of edition_date rows in All-editions
'Find the firstRow
row_runner = twoRow
Do While row_runner <= Final5Row
    bn = WCdbIM.Cells(row_runner, d_col)
    If WCdbIM.Cells(row_runner, d_col) = edition_date Then
    startRow = row_runner
    twoRow = startRow
    row_runner = Final5Row + 10
    End If
row_runner = row_runner + 1
Loop
'Find the last Row
row_runner = twoRow
endRow = Final5Row
Do While row_runner <= Final5Row
    zn = WCdbIM.Cells(row_runner, d_col)
    If WCdbIM.Cells(row_runner, d_col) <> edition_date Then
    endRow = row_runner - 1
    row_runner = Final5Row + 10
    End If
row_runner = row_runner + 1
Loop
WLIM.Range(WLIM.Cells(sixRow, 1), WLIM.Cells(endRow - startRow + 2, Final5Col)).value = WCdbIM.Range(WCdbIM.Cells(startRow, 1), WCdbIM.Cells(endRow, Final5Col)).value
'Set the password
WLIM.Protect Password

'First Editions Code comparessment
   'Find the codes
oneRow = name1Row + 1
Do While oneRow <= Final1Row
 i = 3
 If WC.Cells(oneRow, first_date_col) = "" Then
        WC.Cells(oneRow, first_col) = "4-Not banned"
 Else
  Do While i < 12
    If WC.Cells(oneRow, i) = WC.Cells(oneRow, first_date_col) Then
        If WC.Cells(name1Row, i) = "XXX" Or WC.Cells(name1Row, i) = "XXXX" Or WC.Cells(name1Row, i) = "XXXX-XX" Or WC.Cells(name1Row, i) = "XX" Or WC.Cells(name1Row, i) = "XXXX-X" Then
        WC.Cells(oneRow, first_col) = "1-Banned"
        Exit Do
        ElseIf WC.Cells(name1Row, i) = "XXXX-0000" Or WC.Cells(name1Row, i) = "XXXX-XX-00" Or WC.Cells(name1Row, i) = "XXXX-XXXX" Then
        WC.Cells(oneRow, first_col) = "2-Likely banned"
        Exit Do
        ElseIf WC.Cells(name1Row, i) = "XXXX-XXX" Then
        WC.Cells(oneRow, first_col) = "3-Undefined"
        Exit Do
        End If
    End If
  i = i + 1
  Loop
 End If
oneRow = oneRow + 1
Loop

Debug.Print "Reached point: Last Date"

'Last Editions Code comparessment
   'Find the codes
oneRow = name1Row + 1
Do While oneRow <= Final1Row
 i = 17
 If WC.Cells(oneRow, last_date_col) = "" Then
        WC.Cells(oneRow, last_col) = "4-Not banned"
 Else
  Do While i < 26
    If WC.Cells(oneRow, i) = WC.Cells(oneRow, last_date_col) Then
        If WC.Cells(name1Row, i) = "XXXX" Then
        form = Left(WC.Cells(oneRow, 2), 4)
        WC.Cells(oneRow, last_col) = "1-Banned"
        Formula1_1 = "=IFERROR(INDEX(Last_edition_import!$C$" & sixRow & ":$C$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_annex).Formula = Formula1_1
        Formula1_2 = "=IFERROR(INDEX(Last_edition_import!$D$" & sixRow & ":$D$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_article).Formula = Formula1_2
        Exit Do
        ElseIf WC.Cells(name1Row, i) = "XXXX-XX" Then
        form = Left(WC.Cells(oneRow, 2), 6)
        WC.Cells(oneRow, last_col) = "1-Banned"
        Formula1_1 = "=IFERROR(INDEX(Last_edition_import!$C$" & sixRow & ":$C$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_annex).Formula = Formula1_1
        Formula1_2 = "=IFERROR(INDEX(Last_edition_import!$D$" & sixRow & ":$D$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_article).Formula = Formula1_2
        Exit Do
        ElseIf WC.Cells(name1Row, i) = "XXXX-X" Then
        form = Left(WC.Cells(oneRow, 2), 5)
        WC.Cells(oneRow, last_col) = "1-Banned"
        Formula1_1 = "=IFERROR(INDEX(Last_edition_import!$C$" & sixRow & ":$C$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_annex).Formula = Formula1_1
        Formula1_2 = "=IFERROR(INDEX(Last_edition_import!$D$" & sixRow & ":$D$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_article).Formula = Formula1_2
        Exit Do
        ElseIf WC.Cells(name1Row, i) = "XX" Then
        form = Left(WC.Cells(oneRow, 2), 2)
        WC.Cells(oneRow, last_col) = "1-Banned"
        Formula1_1 = "=IFERROR(INDEX(Last_edition_import!$C$" & sixRow & ":$C$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_annex).Formula = Formula1_1
        Formula1_2 = "=IFERROR(INDEX(Last_edition_import!$D$" & sixRow & ":$D$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_article).Formula = Formula1_2
        Exit Do
        ElseIf WC.Cells(name1Row, i) = "XXX" Then
        form = Left(WC.Cells(oneRow, 2), 3)
        WC.Cells(oneRow, last_col) = "1-Banned"
        Formula1_1 = "=IFERROR(INDEX(Last_edition_import!$C$" & sixRow & ":$C$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_annex).Formula = Formula1_1
        Formula1_2 = "=IFERROR(INDEX(Last_edition_import!$D$" & sixRow & ":$D$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_article).Formula = Formula1_2
        Exit Do
        
        ElseIf WC.Cells(name1Row, i) = "XXXX-0000" Then
        WC.Cells(oneRow, last_col) = "2-Likely banned"
        form = Val(Left(WC.Cells(oneRow, 2), 4) & "0000")
        Formula1_1 = "=IFERROR(INDEX(Last_edition_import!$C$" & sixRow & ":$C$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_annex).Formula = Formula1_1
        Formula1_2 = "=IFERROR(INDEX(Last_edition_import!$D$" & sixRow & ":$D$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_article).Formula = Formula1_2
        Exit Do
        ElseIf WC.Cells(name1Row, i) = "XXXX-XX-00" Then
        WC.Cells(oneRow, last_col) = "2-Likely banned"
        form = Val(Left(WC.Cells(oneRow, 2), 6) & "00")
        Formula1_1 = "=IFERROR(INDEX(Last_edition_import!$C$" & sixRow & ":$C$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_annex).Formula = Formula1_1
        Formula1_2 = "=IFERROR(INDEX(Last_edition_import!$D$" & sixRow & ":$D$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_article).Formula = Formula1_2
        Exit Do
        ElseIf WC.Cells(name1Row, i) = "XXXX-XXXX" Then
        WC.Cells(oneRow, last_col) = "2-Likely banned"
        form = Left(WC.Cells(oneRow, 2), 8)
        Formula1_1 = "=IFERROR(INDEX(Last_edition_import!$C$" & sixRow & ":$C$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_annex).Formula = Formula1_1
        Formula1_2 = "=IFERROR(INDEX(Last_edition_import!$D$" & sixRow & ":$D$" & Final6Row & ",MATCH(" & form & ",Last_edition_import!$A$" & sixRow & ":$A$" & Final6Row & ",0)),"""")"
        WC.Cells(oneRow, last_article).Formula = Formula1_2
        Exit Do
        
        ElseIf WC.Cells(name1Row, i) = "XXXX-XXX" Then
        WC.Cells(oneRow, last_col) = "3-Undefined"
        Exit Do
        End If
    End If
  i = i + 1
  Loop
 End If
oneRow = oneRow + 1
Loop

'Copy the result into Main sheet
oneRow = 3
' Get range for column 2 in WC
    Set valueRangeWM = WM.Cells(fourRow, code_col).Resize(Final4Row - name4Row, 1)
    
    ' Get range for column 3 in WC and convert to array
    Set valueRangeWCdb = WC.Cells(oneRow, 2).Resize(Final1Row - 2, 1)
    valueArrWCdb = valueRangeWCdb.value
    
    ' Loop through each value in column 2 of WC
    For i = 1 To valueRangeWM.Rows.Count
        matchFound = False
        
        ' Loop through each value in array from column 3 of WCdb
        'Check if there is more then one HS code in array
        If IsArray(valueArrWCdb) Then
        For j = LBound(valueArrWCdb) To UBound(valueArrWCdb)
            xz = LBound(valueArrWCdb)
            zx = UBound(valueArrWCdb)
            x = valueArrWCdb(j, 1)
            y = valueRangeWM.Cells(i, 1)
            If valueArrWCdb(j, 1) = valueRangeWM.Cells(i, 1).value Then
                ' Match found
                matchFound = True
                Exit For
            End If
        Next j
        Else
        j = 1
        End If
        WM.Cells(i + name4Row, f_date_col) = WC.Cells(j + 2, first_col)
        WM.Cells(i + name4Row, f_date_col + 1) = WC.Cells(j + 2, first_date_col)
        WM.Cells(i + name4Row, l_date_col) = WC.Cells(j + 2, last_col)
        WM.Cells(i + name4Row, l_annex) = WC.Cells(j + 2, last_annex)
        WM.Cells(i + name4Row, l_article) = WC.Cells(j + 2, last_article)
        WM.Cells(i + name4Row, l_date_col + 1) = edition_date
        If WM.Cells(i + name4Row, f_date_col) = "1-Banned" And WM.Cells(i + name4Row, f_date_col) <> WM.Cells(i + name4Row, l_date_col) Then
        WM.Cells(i + name4Row, l_date_col).Font.ColorIndex = 3
        End If
        
        Next i
        
Debug.Print "Reached point: Transaction's Date"
WC.Protect Password

'Stop
Call Code_transactions
Call Code_extra
        
' Enable screen updating after the macro is done
WM.Activate
Application.ScreenUpdating = True
    
'MsgBox ("Macros is finished. Check the results")
End Sub

