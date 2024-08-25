Attribute VB_Name = "codes_weapons_priority"
Sub Code_extra()

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
'Set WC = Worksheets("Codes_transaction")
Set WC = Worksheets("Codes_hp")
Set WCW = Worksheets("Codes_weapon")
Set WM = Worksheets("Main")
Set WHP = Worksheets("High_Priority")
Set WW = Worksheets("Weapons")

oneRow = 3
twoRow = 3
threeRow = 3
fourRow = 4
fiveRow = 3
name1Row = oneRow - 1
name2Row = twoRow - 1
name3Row = threeRow - 1
name4Row = fourRow - 1
name5Row = fiveRow - 1

Final1Col = WC.Range("IV2").End(xlToLeft).Column
Final1Row = WC.Range("B65536").End(xlUp).Row
Final2Row = WHP.Range("A65536").End(xlUp).Row
Final2Col = WHP.Range("IV1").End(xlToLeft).Column
Final3Row = WW.Range("A65536").End(xlUp).Row
Final3Col = WW.Range("IV1").End(xlToLeft).Column
Final4Row = WM.Range("A65536").End(xlUp).Row
Final4Col = WM.Range("IV2").End(xlToLeft).Column
Final5Row = WCW.Range("B65536").End(xlUp).Row
Final5Col = WCW.Range("IV2").End(xlToLeft).Column

'Find columns in Codes_hp
first_col = WorksheetFunction.Match("High-priority Items (last edition) (Yes, No)", WC.Rows(name1Row), 0)
ban_col1 = first_col - 2

'Find columns in Codes_weapons
second_col = WorksheetFunction.Match("Weapon (Yes, No)", WCW.Rows(name5Row), 0)
ban_col2 = second_col - 2


'Find columns in Main
code_col = WorksheetFunction.Match("HS Code", WM.Rows(name4Row), 0)
'date_trans_col = WorksheetFunction.Match("Date", WM.Rows(name1Row), 0)
'f_date_col = WorksheetFunction.Match("First Editions Result", WM.Rows(name1Row), 0)
't_date_col = WorksheetFunction.Match("Transaction's date Result (Grace period is ignored)", WM.Rows(name1Row), 0)
'editions_date_col = t_date_col + 1
t_annex_col = WorksheetFunction.Match("Transaction's_Annex", WM.Rows(name4Row), 0)
t_article_col = WorksheetFunction.Match("Transaction's_Article", WM.Rows(name4Row), 0)
'grace_col = WorksheetFunction.Match("Transaction's date Result with grace", WM.Rows(name4Row), 0)
'day1_col = WorksheetFunction.Match("Days since 1-Banned", WM.Rows(name4Row), 0)
'day2_col = WorksheetFunction.Match("Days since 2-Likely banned", WM.Rows(name4Row), 0)
'date1bann_col = WorksheetFunction.Match("Date first edition when BAN", WM.Rows(name4Row), 0)
high_priority_col = WorksheetFunction.Match("High-priority Items (last edition) (Yes, No)", WM.Rows(name4Row), 0)
weapon_col = WorksheetFunction.Match("Weapon  (Yes, No)", WM.Rows(name4Row), 0)
transit_col = WorksheetFunction.Match("Transit prohibited (last edition) (Yes, No)", WM.Rows(name4Row), 0)
l_annex = WorksheetFunction.Match("Last Edition Annex", WM.Rows(name4Row), 0)
l_article = l_annex + 1
' Set the password for the protected sheet
Password = "U82024"
' Reference the protected worksheet
WC.Unprotect Password
WCW.Unprotect Password


'GoTo Test_Last

'Clean the prev. results from Codes
WC.Range(WC.Cells(oneRow, 3), WC.Cells(oneRow + 3000, Final1Col)).Font.ColorIndex = xlAutomatic
WC.Range(WC.Cells(oneRow, 2), WC.Cells(oneRow + 3000, 2)).ClearContents 'Del HS Code
WC.Range(WC.Cells(oneRow, first_col), WC.Cells(oneRow + 3000, first_col)).ClearContents
WCW.Range(WCW.Cells(fiveRow, 3), WCW.Cells(fiveRow + 3000, Final5Col)).Font.ColorIndex = xlAutomatic
WCW.Range(WCW.Cells(fiveRow, 2), WCW.Cells(fiveRow + 3000, 2)).ClearContents 'Del HS Code
WCW.Range(WCW.Cells(fiveRow, second_col), WCW.Cells(fiveRow + 3000, second_col)).ClearContents

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
' Create a HS Code (Copy from HS Code WC sheet)
Set CopyRange = WC.Range(WC.Cells(oneRow, 2), WC.Cells(Final1Row, 2))
Set PastRange = WCW.Range(WCW.Cells(fiveRow, 2), WCW.Cells(Final1Row, 2))
PastRange.value = CopyRange.value
Final5Col = WCW.Range("IV2").End(xlToLeft).Column
Final5Row = WCW.Range("B65536").End(xlUp).Row

Test_Last:
'Codes_hp comparessment
oneRow = name1Row + 1
Do While oneRow <= Final1Row
 i = 3
 If WC.Cells(oneRow, ban_col1) = 0 Then
        WC.Cells(oneRow, first_col) = "No"
 Else
  Do While i < 10
    If WC.Cells(oneRow, i) = 1 Then
        If WC.Cells(name1Row, i) = "XXX" Or WC.Cells(name1Row, i) = "XXXX" Or WC.Cells(name1Row, i) = "XXXX-XX" Or WC.Cells(name1Row, i) = "XX" Or WC.Cells(name1Row, i) = "XXXX-X" Then
        WC.Cells(oneRow, first_col) = "Yes"
        Exit Do
        ElseIf WC.Cells(name1Row, i) = "XXXX-0000" Or WC.Cells(name1Row, i) = "XXXX-XX-00" Or WC.Cells(name1Row, i) = "XXXX-XXXX" Then
        WC.Cells(oneRow, first_col) = "Likely Yes"
        Exit Do
        ElseIf WC.Cells(name1Row, i) = "XXXX-XXX" Then
        WC.Cells(oneRow, first_col) = "Undefined"
        Exit Do
        End If
    End If
  i = i + 1
  Loop
 End If
oneRow = oneRow + 1
Loop

'Codes_weapons comparessment
fiveRow = name5Row + 1
Do While fiveRow <= Final5Row
 i = 3
 If WCW.Cells(fiveRow, ban_col2) = 0 Then
        WCW.Cells(fiveRow, second_col) = "No"
 Else
  Do While i < 10
    If WCW.Cells(fiveRow, i) = 1 Then
        If WCW.Cells(name1Row, i) = "XXX" Or WCW.Cells(name5Row, i) = "XXXX" Or WCW.Cells(name5Row, i) = "XXXX-XX" Or WCW.Cells(name5Row, i) = "XX" Or WCW.Cells(name5Row, i) = "XXXX-X" Then
        WCW.Cells(fiveRow, second_col) = "Yes"
        Exit Do
        ElseIf WCW.Cells(name5Row, i) = "XXXX-0000" Or WCW.Cells(name5Row, i) = "XXXX-XX-00" Or WCW.Cells(name5Row, i) = "XXXX-XXXX" Then
        WCW.Cells(fiveRow, second_col) = "Likely Yes"
        Exit Do
        ElseIf WCW.Cells(name5Row, i) = "XXXX-XXX" Then
        WCW.Cells(fiveRow, second_col) = "Undefined"
        Exit Do
        End If
    End If
  i = i + 1
  Loop
 End If
fiveRow = fiveRow + 1
Loop

'Copy the result into Main sheet from WC
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
        If IsArray(valueArrWCdb) Then
        For j = LBound(valueArrWCdb) To UBound(valueArrWCdb)
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
        WM.Cells(i + name4Row, high_priority_col) = WC.Cells(j + 2, first_col)
       
        Next i
'Copy the result into Main sheet from WCW
fiveRow = 3
' Get range for column 2 in WCW
    Set valueRangeWM = WM.Cells(fourRow, code_col).Resize(Final4Row - name4Row, 1)
    
    ' Get range for column 3 in WC and convert to array
    Set valueRangeWCdb = WCW.Cells(fiveRow, 2).Resize(Final5Row - 2, 1)
    valueArrWCdb = valueRangeWCdb.value
    
    ' Loop through each value in column 2 of WC
    For i = 1 To valueRangeWM.Rows.Count
        matchFound = False
        
        ' Loop through each value in array from column 3 of WCdb
        If IsArray(valueArrWCdb) Then
        For j = LBound(valueArrWCdb) To UBound(valueArrWCdb)
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
        WM.Cells(i + name4Row, weapon_col) = WCW.Cells(j + 2, second_col)
       
        Next i
        
'Fill column Transit
fourRow = name4Row + 1
Do While fourRow <= Final4Row
    If WM.Cells(fourRow, l_annex) = "ANNEX VII" Or WM.Cells(fourRow, l_annex) = "ANNEX XI" Or WM.Cells(fourRow, l_annex) = "ANNEX XX" Or WM.Cells(fourRow, l_annex) = "ANNEX XXXV" Or WM.Cells(fourRow, l_annex) = "ANNEX XXXVII" Then
    WM.Cells(fourRow, transit_col) = "Yes"
    Else
    WM.Cells(fourRow, transit_col) = "No"
    End If
fourRow = fourRow + 1
Loop


WC.Protect Password
WCW.Protect Password



End Sub

