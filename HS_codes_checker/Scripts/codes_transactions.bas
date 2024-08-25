Attribute VB_Name = "codes_transactions"
Sub Code_transactions()

Dim objSell As Excel.Workbook, objShab As Excel.Workbook
Dim ws As Worksheet
Dim WErr As Worksheet
Dim PT As PivotTable
Dim text As String
Dim Final1Row As Integer
Dim Final1Col As Integer
Dim NowDate As Long
Dim DateDiff As Integer
Dim dateValue As Date

Set objThis = Excel.ActiveWorkbook
Set WC = Worksheets("Codes_transaction")
Set WCdb = Worksheets("All_editions")
Set WE = Worksheets("Editions")
Set WM = Worksheets("Main")
Set WCdbIM = Worksheets("All_editions_import") 'Must be filled manualy before (if needed)
Set WLIM = Worksheets("Last_edition_import") 'Changes automaticly
Set WCC = Worksheets("Codes_first_last")

oneRow = 3
twoRow = 2
threeRow = 2
fourRow = 4
fiveRow = 2
sixRow = 2
sevenRow = 3
name1Row = oneRow - 1
name2Row = twoRow - 1
name3Row = threeRow - 1
name4Row = fourRow - 1
name5Row = fiveRow - 1
name6Row = sixRow - 1
name7Row = sevenRow - 1

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
Final6Row = WLIM.Range("A65536").End(xlUp).Row
Final6Col = WLIM.Range("IV1").End(xlToLeft).Column
Final7Col = WCC.Range("IV2").End(xlToLeft).Column
Final7Row = WCC.Range("B65536").End(xlUp).Row

'Find columns in All_editions
c_col = WorksheetFunction.Match("CN", WCdb.Rows(name2Row), 0)
d_col = WorksheetFunction.Match("Date_of_publication", WCdb.Rows(name2Row), 0)
import_col = WorksheetFunction.Match("Import/Export", WCdb.Rows(name2Row), 0)
annex_col = WorksheetFunction.Match("Annex", WCdb.Rows(name2Row), 0)
article_col = WorksheetFunction.Match("Article", WCdb.Rows(name2Row), 0)

'Find columns in Codes
next_digit_col1 = WorksheetFunction.Match("Next_digit", WC.Rows(name1Row), 0)
first_col = WorksheetFunction.Match("Transaction's date Result", WC.Rows(name1Row), 0)
ban_col1 = first_col - 1
first_date_col = first_col + 1
x2_col = WorksheetFunction.Match("XX", WC.Rows(name1Row), 0)
x8_col = WorksheetFunction.Match("XXXX-XXXX", WC.Rows(name1Row), 0)
trans_annex_col = WorksheetFunction.Match("Transaction's_Annex", WC.Rows(name1Row), 0)
trans_article_col = trans_annex_col + 1
trans_grace_col = trans_article_col + 1

'Find columns in Editions
date_edition_col = WorksheetFunction.Match("Edition's date", WE.Rows(name3Row), 0)
start_col = WorksheetFunction.Match("Start_Row", WE.Rows(name3Row), 0)

'Find columns in Codes first last
code_col1 = WorksheetFunction.Match("HS Code", WCC.Rows(name7Row), 0)
x4_col = WorksheetFunction.Match("XXXX", WCC.Rows(name7Row), 0)
x6_col = WorksheetFunction.Match("XXXX-XX", WCC.Rows(name7Row), 0)

'Find columns in Main
code_col = WorksheetFunction.Match("HS Code", WM.Rows(name4Row), 0)
date_trans_col = WorksheetFunction.Match("Date", WM.Rows(name4Row), 0)
f_date_col = WorksheetFunction.Match("First Editions Result", WM.Rows(name4Row), 0)
t_date_col = WorksheetFunction.Match("Transaction's date Result (Grace period is ignored)", WM.Rows(name4Row), 0)
editions_date_col = t_date_col + 1
t_annex_col = WorksheetFunction.Match("Transaction's_Annex", WM.Rows(name4Row), 0)
t_article_col = WorksheetFunction.Match("Transaction's_Article", WM.Rows(name4Row), 0)
grace_col = WorksheetFunction.Match("Transaction's date Result with grace", WM.Rows(name4Row), 0)
day1_col = WorksheetFunction.Match("Days since 1-Banned", WM.Rows(name4Row), 0)
day2_col = WorksheetFunction.Match("Days since 2-Likely banned", WM.Rows(name4Row), 0)
date1bann_col = WorksheetFunction.Match("Date first edition when BAN", WM.Rows(name4Row), 0)
high_priority_col = WorksheetFunction.Match("High-priority Items (last edition) (Yes, No)", WM.Rows(name4Row), 0)
weapon_col = WorksheetFunction.Match("Weapon  (Yes, No)", WM.Rows(name4Row), 0)
transit_col = WorksheetFunction.Match("Transit prohibited (last edition) (Yes, No)", WM.Rows(name4Row), 0)
grace_date_col = WorksheetFunction.Match("Grace Period", WM.Rows(name4Row), 0)


' Set the password for the protected sheet
Password = "U82024" ' Replace with the actual password
' Reference the protected worksheet
WC.Unprotect Password

Set HSRangeWCdb = WCdb.Cells(twoRow, c_col).Resize(Final2Row - 2, d_col)
RowCount = HSRangeWCdb.Rows.Count
ColCount = HSRangeWCdb.Columns.Count

'GoTo N28

'Clean the prev. results from Codes
WC.Range(WC.Cells(oneRow, 3), WC.Cells(oneRow + 64000, Final1Col)).Font.ColorIndex = xlAutomatic
WC.Range(WC.Cells(oneRow, 1), WC.Cells(oneRow + 64000, 2)).ClearContents 'Del HS Code
WC.Range(WC.Cells(oneRow, x2_col), WC.Cells(oneRow + 64000, x8_col)).ClearContents
WC.Range(WC.Cells(oneRow, first_col), WC.Cells(oneRow + 64000, first_col)).ClearContents
WC.Range(WC.Cells(oneRow, trans_annex_col), WC.Cells(oneRow + 64000, trans_grace_col)).ClearContents

' Convert the WM. date_trans_col into Date (to avoid missgtakes in comparing and calculations alter
' Loop through each cell in the column
    For i = fourRow To Final4Row
        ' Get the date string from the cell
        dateString = WM.Cells(i, date_trans_col).value
        If dateString <> "" Then
            ' Convert the date string to a Date data type
            dateValue = CDate(dateString)
            ' Write the converted date back to the cell
            WM.Cells(i, date_trans_col).value = dateValue
            ' Format the cell as a date
            WM.Cells(i, date_trans_col).NumberFormat = "dd.mm.yyyy"
        End If
    Next i

' Create a HS Code (Copy from HS Code WM sheet)
Set CopyRange = WM.Range(WM.Cells(fourRow, 1), WM.Cells(Final4Row, 2))
Set PastRange = WC.Range(WC.Cells(oneRow, 1), WC.Cells(Final4Row - 1, 2))
PastRange.value = CopyRange.value
Final1Col = WC.Range("IV2").End(xlToLeft).Column
Final1Row = WC.Range("B65536").End(xlUp).Row

'Create row frames for every edition
WE.Activate
threeRow = 2
fiveRow = 2
WE.Cells(threeRow, start_col) = fiveRow
threeRow = threeRow + 1
Do While threeRow <= Final3Row
 edition_date = WE.Cells(threeRow, date_edition_col)
 previousDate = WCdbIM.Cells(fiveRow, d_col).value
      Do While fiveRow <= Final5Row
        currentDate = WCdbIM.Cells(fiveRow, d_col).value
         ' Check if a new month has started
        If currentDate <> previousDate Then
            ' Record the start row for the current month
            WE.Cells(threeRow, start_col) = fiveRow
            Exit Do
        End If
        previousDate = currentDate
    fiveRow = fiveRow + 1
    Loop

threeRow = threeRow + 1
Loop

'Stop
   
'Editions Code comparessment
    'Find a right date for last edition
'Test_Last:

Do While oneRow <= Final1Row
    twoRow = 2
    threeRow = 2
 If WC.Cells(oneRow, 1) < WE.Cells(threeRow, date_edition_col) Then
        'WC.Cells(oneRow, first_date_col) = "Before 1st edition"
        WC.Cells(oneRow, first_col) = "4-Not banned"
        GoTo M3
    
 ElseIf WC.Cells(oneRow, 1) <> WC.Cells(oneRow - 1, 1) Then
    WM.Activate
    trans_date = WC.Cells(oneRow, 1)
    WE.Activate
    '
    threeRow = 2
    Do While threeRow <= Final3Row
        If WE.Cells(threeRow, date_edition_col) < trans_date Then
        edition_date = WE.Cells(threeRow, date_edition_col)
        startRow = WE.Cells(threeRow, start_col)
            If threeRow = Final3Row Then
            endRow = Final5Row
            Else
            endRow = WE.Cells(threeRow + 1, start_col)
            End If
        Else
            Exit Do
        End If
    threeRow = threeRow + 1
    Loop
    
 End If

'Fill the formulas in XXX columns

'XXXX-XXX
form = Left(WC.Cells(oneRow, 2), 7)
Formula = "=IFERROR(INDEX(All_editions_import!$H$" & startRow & ":$H$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
WC.Cells(oneRow, x2_col + 7).Formula = Formula
If WC.Cells(oneRow, x2_col + 7) <> "" Then
    Formula1_1 = "=IFERROR(INDEX(All_editions_import!$C$" & startRow & ":$C$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_annex_col).Formula = Formula1_1
    Formula1_2 = "=IFERROR(INDEX(All_editions_import!$D$" & startRow & ":$D$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_article_col).Formula = Formula1_2
    Formula1_3 = "=IFERROR(INDEX(All_editions_import!$I$" & startRow & ":$I$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_grace_col).Formula = Formula1_3
    If WC.Cells(oneRow, trans_grace_col) = 0 Then WC.Cells(oneRow, trans_grace_col).Delete
End If

'XXXX-XXXX
form = Left(WC.Cells(oneRow, 2), 8)
Formula = "=IFERROR(INDEX(All_editions_import!$H$" & startRow & ":$H$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
WC.Cells(oneRow, x8_col).Formula = Formula
If WC.Cells(oneRow, x8_col) <> "" Then
    Formula1_1 = "=IFERROR(INDEX(All_editions_import!$C$" & startRow & ":$C$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_annex_col).Formula = Formula1_1
    Formula1_2 = "=IFERROR(INDEX(All_editions_import!$D$" & startRow & ":$D$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_article_col).Formula = Formula1_2
    Formula1_3 = "=IFERROR(INDEX(All_editions_import!$I$" & startRow & ":$I$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_grace_col).Formula = Formula1_3
    If WC.Cells(oneRow, trans_grace_col) = 0 Then WC.Cells(oneRow, trans_grace_col).Delete
End If

'XXXX-XX-00
form = Val(Left(WC.Cells(oneRow, 2), 6) & "00")
Formula = "=IFERROR(INDEX(All_editions_import!$H$" & startRow & ":$H$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
WC.Cells(oneRow, x2_col + 6).Formula = Formula
If WC.Cells(oneRow, x2_col + 6) <> "" Then
    Formula1_1 = "=IFERROR(INDEX(All_editions_import!$C$" & startRow & ":$C$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_annex_col).Formula = Formula1_1
    Formula1_2 = "=IFERROR(INDEX(All_editions_import!$D$" & startRow & ":$D$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_article_col).Formula = Formula1_2
    Formula1_3 = "=IFERROR(INDEX(All_editions_import!$I$" & startRow & ":$I$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_grace_col).Formula = Formula1_3
    If WC.Cells(oneRow, trans_grace_col) = 0 Then WC.Cells(oneRow, trans_grace_col).Delete
End If

'XXXX-0000
form = Val(Left(WC.Cells(oneRow, 2), 4) & "0000")
Formula = "=IFERROR(INDEX(All_editions_import!$H$" & startRow & ":$H$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
WC.Cells(oneRow, x2_col + 5).Formula = Formula
If WC.Cells(oneRow, x2_col + 5) <> "" Then
    Formula1_1 = "=IFERROR(INDEX(All_editions_import!$C$" & startRow & ":$C$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_annex_col).Formula = Formula1_1
    Formula1_2 = "=IFERROR(INDEX(All_editions_import!$D$" & startRow & ":$D$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_article_col).Formula = Formula1_2
    Formula1_3 = "=IFERROR(INDEX(All_editions_import!$I$" & startRow & ":$I$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_grace_col).Formula = Formula1_3
    If WC.Cells(oneRow, trans_grace_col) = 0 Then WC.Cells(oneRow, trans_grace_col).Delete
End If

'XX
form = Left(WC.Cells(oneRow, 2), 2)
Formula = "=IFERROR(INDEX(All_editions_import!$H$" & startRow & ":$H$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
WC.Cells(oneRow, x2_col).Formula = Formula
If WC.Cells(oneRow, x2_col) <> "" Then
    Formula1_1 = "=IFERROR(INDEX(All_editions_import!$C$" & startRow & ":$C$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_annex_col).Formula = Formula1_1
    Formula1_2 = "=IFERROR(INDEX(All_editions_import!$D$" & startRow & ":$D$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_article_col).Formula = Formula1_2
    Formula1_3 = "=IFERROR(INDEX(All_editions_import!$I$" & startRow & ":$I$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_grace_col).Formula = Formula1_3
    If WC.Cells(oneRow, trans_grace_col) = 0 Then WC.Cells(oneRow, trans_grace_col).Delete
End If

'XXX
form = Left(WC.Cells(oneRow, 2), 3)
Formula = "=IFERROR(INDEX(All_editions_import!$H$" & startRow & ":$H$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
WC.Cells(oneRow, x2_col + 1).Formula = Formula
If WC.Cells(oneRow, x2_col + 1) <> "" Then
    Formula1_1 = "=IFERROR(INDEX(All_editions_import!$C$" & startRow & ":$C$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_annex_col).Formula = Formula1_1
    Formula1_2 = "=IFERROR(INDEX(All_editions_import!$D$" & startRow & ":$D$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_article_col).Formula = Formula1_2
    Formula1_3 = "=IFERROR(INDEX(All_editions_import!$I$" & startRow & ":$I$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_grace_col).Formula = Formula1_3
    If WC.Cells(oneRow, trans_grace_col) = 0 Then WC.Cells(oneRow, trans_grace_col).Delete
End If

'XXXX
form = Left(WC.Cells(oneRow, 2), 4)
Formula = "=IFERROR(INDEX(All_editions_import!$H$" & startRow & ":$H$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
WC.Cells(oneRow, x2_col + 2).Formula = Formula
If WC.Cells(oneRow, x2_col + 2) <> "" Then
    Formula1_1 = "=IFERROR(INDEX(All_editions_import!$C$" & startRow & ":$C$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_annex_col).Formula = Formula1_1
    Formula1_2 = "=IFERROR(INDEX(All_editions_import!$D$" & startRow & ":$D$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_article_col).Formula = Formula1_2
    Formula1_3 = "=IFERROR(INDEX(All_editions_import!$I$" & startRow & ":$I$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_grace_col).Formula = Formula1_3
    If WC.Cells(oneRow, trans_grace_col) = 0 Then WC.Cells(oneRow, trans_grace_col).Delete
End If

'XXXX-X
form = Left(WC.Cells(oneRow, 2), 5)
Formula = "=IFERROR(INDEX(All_editions_import!$H$" & startRow & ":$H$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
WC.Cells(oneRow, x2_col + 3).Formula = Formula
If WC.Cells(oneRow, x2_col + 3) <> "" Then
    Formula1_1 = "=IFERROR(INDEX(All_editions_import!$C$" & startRow & ":$C$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_annex_col).Formula = Formula1_1
    Formula1_2 = "=IFERROR(INDEX(All_editions_import!$D$" & startRow & ":$D$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_article_col).Formula = Formula1_2
    Formula1_3 = "=IFERROR(INDEX(All_editions_import!$I$" & startRow & ":$I$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_grace_col).Formula = Formula1_3
    If WC.Cells(oneRow, trans_grace_col) = 0 Then WC.Cells(oneRow, trans_grace_col).Delete
End If

'XXXX-XX
form = Left(WC.Cells(oneRow, 2), 6)
Formula = "=IFERROR(INDEX(All_editions_import!$H$" & startRow & ":$H$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
WC.Cells(oneRow, x2_col + 4).Formula = Formula
If WC.Cells(oneRow, x2_col + 4) <> "" Then
    Formula1_1 = "=IFERROR(INDEX(All_editions_import!$C$" & startRow & ":$C$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_annex_col).Formula = Formula1_1
    Formula1_2 = "=IFERROR(INDEX(All_editions_import!$D$" & startRow & ":$D$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_article_col).Formula = Formula1_2
    Formula1_3 = "=IFERROR(INDEX(All_editions_import!$I$" & startRow & ":$I$" & endRow & ",MATCH(" & form & ",All_editions_import!$A$" & startRow & ":$A$" & endRow & ",0)),"""")"
    WC.Cells(oneRow, trans_grace_col).Formula = Formula1_3
    If WC.Cells(oneRow, trans_grace_col) = 0 Then WC.Cells(oneRow, trans_grace_col).Delete
End If

M3:
oneRow = oneRow + 1
Loop

'Fill Ban marks for Transactions
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


'### Copy the result into Main sheet

oneRow = 3
fourRow = name4Row + 1
' Create a HS Code (Copy from HS Code WM sheet)
Set CopytRange = WC.Range(WC.Cells(oneRow, first_col), WC.Cells(Final1Row, trans_grace_col))
Set PastRange = WM.Range(WM.Cells(fourRow, t_date_col), WM.Cells(Final4Row, grace_date_col))
PastRange.value = CopytRange.value

WC.Protect Password
WM.Activate


'Grace period
'Days in Bann
grace_date = 0
fourRow = name4Row + 1
Do While fourRow <= Final4Row

If WM.Cells(fourRow, f_date_col + 1) >= WM.Cells(fourRow, date_trans_col) Then
    WM.Cells(fourRow, grace_col) = "4-Not banned"

Else

' 4 and 4
If WM.Cells(fourRow, f_date_col) = "4-Not banned" And WM.Cells(fourRow, f_date_col) = WM.Cells(fourRow, t_date_col) Then
    WM.Cells(fourRow, grace_col) = WM.Cells(fourRow, f_date_col)

'1 and 1
ElseIf WM.Cells(fourRow, f_date_col) = "1-Banned" And WM.Cells(fourRow, f_date_col) = WM.Cells(fourRow, t_date_col) Then
    If WM.Cells(fourRow, grace_date_col) = "" Then
    WM.Cells(fourRow, day1_col) = WM.Cells(fourRow, date_trans_col) - WM.Cells(fourRow, f_date_col + 1)
    grace_date = 0
    Else
    WM.Cells(fourRow, day1_col) = WM.Cells(fourRow, date_trans_col) - WM.Cells(fourRow, grace_date_col)
    grace_date = WM.Cells(fourRow, grace_date_col)
    End If
    If WM.Cells(fourRow, day1_col) > 0 Then
    WM.Cells(fourRow, grace_col) = WM.Cells(fourRow, f_date_col)
    Else
    WM.Cells(fourRow, grace_col) = "4-Not banned"
    End If

'2 and 2
ElseIf WM.Cells(fourRow, f_date_col) = "2-Likely banned" And WM.Cells(fourRow, f_date_col) = WM.Cells(fourRow, t_date_col) Then
    If WM.Cells(fourRow, grace_date_col) = "" Then
    WM.Cells(fourRow, day2_col) = WM.Cells(fourRow, date_trans_col) - WM.Cells(fourRow, f_date_col + 1)
    grace_date = 0
    Else
    WM.Cells(fourRow, day2_col) = WM.Cells(fourRow, date_trans_col) - WM.Cells(fourRow, grace_date_col)
    grace_date = WM.Cells(fourRow, grace_date_col)
    End If
    If WM.Cells(fourRow, day2_col) > 0 Then
    WM.Cells(fourRow, grace_col) = WM.Cells(fourRow, f_date_col)
    Else
    WM.Cells(fourRow, grace_col) = "4-Not banned"
    End If
    
'2 and 1
ElseIf WM.Cells(fourRow, f_date_col) = "2-Likely banned" And WM.Cells(fourRow, t_date_col) = "1-Banned" Then
    sevenRow = 3
    Do While sevenRow < Final7Row
     If WM.Cells(fourRow, code_col) = WCC.Cells(sevenRow, code_col1) Then
        For i = x4_col - 1 To x6_col
            If WCC.Cells(sevenRow, i) <> "" Then
            WM.Cells(fourRow, day1_col) = WM.Cells(fourRow, date_trans_col) - WCC.Cells(sevenRow, i)
            WM.Cells(fourRow, date1bann_col) = WCC.Cells(sevenRow, i)
            Exit Do
            End If
        Next i
     End If
    sevenRow = sevenRow + 1
    Loop
    If WM.Cells(fourRow, grace_date_col) = "" Then
    WM.Cells(fourRow, day2_col) = WM.Cells(fourRow, date_trans_col) - WM.Cells(fourRow, f_date_col + 1)
    grace_date = 0
    Else
    WM.Cells(fourRow, day2_col) = WM.Cells(fourRow, date_trans_col) - WM.Cells(fourRow, grace_date_col)
    grace_date = WM.Cells(fourRow, grace_date_col)
    End If
    If WM.Cells(fourRow, day2_col) > 0 Then
    WM.Cells(fourRow, grace_col) = WM.Cells(fourRow, t_date_col)
    Else
    WM.Cells(fourRow, grace_col) = WM.Cells(fourRow, f_date_col)
    End If
    
'Was Banned (1 or 2) and became NOT banned
'2 and 4
ElseIf WM.Cells(fourRow, f_date_col) = "2-Likely banned" And WM.Cells(fourRow, t_date_col) = "4-Not banned" Then
    WM.Cells(fourRow, grace_col) = "Attention!!!!"
'1 and NOT 1
ElseIf WM.Cells(fourRow, f_date_col) = "1-Banned" And WM.Cells(fourRow, f_date_col) <> WM.Cells(fourRow, t_date_col) Then
    If WM.Cells(fourRow, f_date_col) > WM.Cells(fourRow, editions_date_col) Then
    WM.Cells(fourRow, grace_col) = WM.Cells(fourRow, t_date_col)
    Else
    WM.Cells(fourRow, grace_col) = "Attention!!!!"
    End If
End If
End If
fourRow = fourRow + 1
Loop



End Sub

