Attribute VB_Name = "f_DellPivotSheets"
Sub f_DellPivotSheets()
    Dim WS As Worksheet
    Dim name1Index As Integer
    
    ' Find the position of the sheet named "Name1"
    For Each WS In Worksheets
        If WS.Name = "Pivots>>" Then
            name1Index = WS.Index
            Exit For
        End If
    Next WS
    
    ' Delete all sheets after the sheet named "Name1"
    Application.DisplayAlerts = False ' Suppress delete confirmation dialog
    For i = Worksheets.Count To name1Index + 1 Step -1
        Worksheets(i).Delete
    Next i
    Application.DisplayAlerts = True
    
'Dell all Hiperlinks
Set WP = Worksheets("Pivots>>")
Set dellrange = WP.Range("A:A")
dellrange.Select
Selection.Hyperlinks.Delete

MsgBox ("All Pivot Tables were deleted")
    
End Sub
