Attribute VB_Name = "codes_import"
Sub codes_import()

Dim objSell As Excel.Workbook, objShab As Excel.Workbook
Dim ws As Worksheet
Dim WErr As Worksheet
Dim PT As PivotTable
Dim text As String
Dim PTCache As PivotCache
Dim PRange As Range, SourseRange As Range, PTDestination As Range
Dim Final1Row As Integer
Dim Final1Col As Integer
'Dim cur As Long
Dim NowDate As Long
Dim DateDiff As Integer

Set objThis = Excel.ActiveWorkbook
Set WC = Worksheets("All_editions")
Set WE = Worksheets("Editions")
oneRow = 2
twoRow = 2
name1Row = oneRow - 1
name2Row = twoRow - 1 'name Row in DB sheet
twoCol = 2

' Set the password for the protected sheet
Password = "U82024" ' Replace with the actual password
' Reference the protected worksheet
WC.Unprotect Password

'Clear the sheet WC
Final1Row = WC.Range("A65536").End(xlUp).Row
Final1Col = WC.Range("IV1").End(xlToLeft).Column
WC.Cells(2, 1).Resize(Final1Row + 10, Final1Col + 5).Clear

Final1Col = WC.Range("IV1").End(xlToLeft).Column
Final1Row = WC.Range("A65536").End(xlUp).Row
Final2Col = WE.Range("IV1").End(xlToLeft).Column
Final2Row = WE.Range("A65536").End(xlUp).Row

'Открываем книгу с базами кодов
    
                 'Open every book with the name in the 1. Row
WE.Activate
Do While twoRow <= Final2Row

    file_name = WE.Cells(twoRow, twoCol)
    Path = "f:\Анализ\UAANALYST\Sources\HS_DB\done_tables\done_tables_v3(Anna)\" & file_name & ".xlsx"
    Set objShab = Excel.Workbooks.Open(Path)
    Set WSS = Worksheets(file_name)
    WSS.Activate
    Final3Row = WSS.Range("A65536").End(xlUp).Row
    Final3Col = WSS.Range("IV1").End(xlToLeft).Column
    Set PRange = WSS.Cells(2, 1).Resize(Final3Row - 1, Final3Col)
    PRange.Copy
    objThis.Activate
    WC.Activate
    Cells(Final1Row + 1, 1).Select
    ActiveSheet.Paste
    Cells(Final1Row + 1, 1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Cells(1, 200).Clear
    objShab.Close (False)
    Final1Row = WC.Range("A65536").End(xlUp).Row

twoRow = twoRow + 1
Loop

WC.Protect Password

End Sub

Sub test()

'Find columns in Code DB
i = 1
c_col = i
Do While WE.Cells(name2Row, i) <> "CN Code"
     i = i + 1
     c_col = i
Loop
'Find columns in Code DB
j = 1
s_col = j
Do While WC.Cells(name1Row, j) <> "Next_digit"
     j = j + 1
     s_col = j
Loop

    'Check the C_rate errors
towRow = 3
oneRow = name1Row + 1
one_col = 2

Do While oneRow <= Final1Row
        i = 4 'numder of digits
        oneCol = 2
    
    If IsEmpty(WC.Cells(oneRow, s_col + 2).value) Then 'check if it already has a date in column Date_Banned
        
        Do While oneCol <= Final1Col - 3
         
            twoRow = 3
            Do While twoRow <= Final2Row
  
              If WC.Cells(oneRow, oneCol) = Left(WE.Cells(twoRow, c_col), i) Then
              WC.Cells(oneRow, oneCol).Font.ColorIndex = 3
              WC.Cells(oneRow, s_col) = Left(WE.Cells(twoRow, c_col), i + 1)
                'check if it is 0 or 1 in the column is_Banned
                   If WC.Cells(oneRow, s_col + 1) = "1" Then WC.Cells(oneRow, s_col + 2) = file_name
              End If
            twoRow = twoRow + 1
            Loop
         
        i = i + 1
        oneCol = oneCol + 1
        Loop
    End If
oneRow = oneRow + 1
Loop
                'Close this book and open a next one
twoCol = twoCol + 1
Loop


MsgBox ("ALL")
End Sub

