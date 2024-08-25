Attribute VB_Name = "save_all_as_pdf"
Sub SaveWorksheetsAsOnePDF()
    Dim savePath As String
    Dim fileName As String
    Dim ws As Worksheet
    Dim selectedSheets As String
    Dim isFirstSheet As Boolean
    
    ' Set the directory path where the PDF file will be saved
    'savePath = "f:\VBProjekt\Projects\VBA_projects\Pivot_tables_2_PDF\PDF tables\" ' Change this to your desired directory
    savePath = ThisWorkbook.Path
    
    ' Construct the file name for the PDF file
    fileName = savePath & "\" & "PDF_tables.pdf"
    
    ' Save all worksheets as one PDF
    'Sheets(Array(Sheets(1).Name)).Select
    'Sheets(Array(Sheets(2).Name)).Select Replace:=True
    'For Each ws In ThisWorkbook.Worksheets
    '    ws.Select False
    'Next ws
    
    ' Initialize isFirstSheet flag
    isFirstSheet = True
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Check if the current worksheet is not the first sheet
        If Not isFirstSheet Then
            ' Select the worksheet (without activating it)
            ws.Select False
        End If
            ' Set isFirstSheet flag to False after processing the first sheet
        isFirstSheet = False
    Next ws
      
    ' Select and activate the first sheet
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fileName, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

MsgBox ("All tables were transfered to PDF. Check the Folder")

End Sub



