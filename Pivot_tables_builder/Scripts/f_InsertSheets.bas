Attribute VB_Name = "f_InsertSheets"
Sub f_InsertNewSheet()

Dim WS As Worksheet
Dim objSell As Excel.Workbook, objShab As Excel.Workbook
Dim WErr As Worksheet
Dim Final1Row As Integer
Dim Final1Col As Integer
Dim text As String
Dim cell As Range


Set objThis = Excel.ActiveWorkbook
Set WP = Worksheets("Pivots>>")
'GoTo M1
oneRow = 4
name1Row = oneRow - 2

Final1Col = WP.Range("IV2").End(xlToLeft).Column
Final1Row = WP.Range("B65536").End(xlUp).Row
    
Do While oneRow <= Final1Row

    ' Insert a new sheet at the end and assign it the name "Name1"
    Set WS = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    WS.Name = WP.Cells(oneRow, 2)
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "'DB-1-B'!R1C1", TextToDisplay:="Back to 'DB-1-B'!"
    Set cell = WP.Cells(oneRow, 1)
    adress = "'" & WS.Name & "'" & "!R1C1"
    text = "'" & oneRow - 2 & "'"
    WP.Hyperlinks.Add Anchor:=cell, Address:="", SubAddress:= _
        adress, TextToDisplay:=text
    
oneRow = oneRow + 1
Loop
M1:

Call p_Pivots_loop

MsgBox ("Pivot tables were built")

End Sub

