Attribute VB_Name = "InsertIndex"
Option Explicit

Sub CreateIndex()
Dim wb As Workbook
Dim wks As Worksheet
Dim sht As Worksheet
Dim i As Integer

Set wb = ThisWorkbook
    
    'Check whether a Index Sheet is already Present
    For i = 1 To wb.Sheets.Count
        If wb.Sheets(i).Name = "Index" Then
            GoTo IndexSht
        End If
    Next i
    
    'Create a new sheet with name "Index"
    wb.Sheets.Add
    wb.Sheets(1).Name = "Index"
 
'Loop Through Each Sheet and get its name
IndexSht:
    Set wks = wb.Sheets("Index")
    wks.Cells(1, 1) = "INDEX"
    wks.Cells(1, 2) = "Number of Rows"
    
    For i = 2 To wb.Sheets.Count
        wks.Cells(i, 1) = wb.Sheets(i).Name
        wks.Cells(i, 2) = wb.Sheets(i).UsedRange.Rows.Count
        'Add hyperlink for easy access to other sheets from Index Sheet
        wks.Hyperlinks.Add Anchor:=wks.Cells(i, 1), Address:="", SubAddress:= _
        "'" & wb.Sheets(i).Name & "'!A1", TextToDisplay:=wb.Sheets(i).Name
    Next i
    
    For Each sht In wb.Worksheets
    
        If sht.Name = "Index" Then
                GoTo NexSht
        End If
        
        If sht.Range("A1").Value = "Back to Index" Then
                GoTo NexCod
        End If
        
        sht.Rows("1:1").Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        sht.Range("A1").FormulaR1C1 = "Back to Index"
        
NexCod:
    
        sht.Range("A1").Hyperlinks.Add Anchor:=sht.Range("A1"), Address:="", SubAddress:= _
            "Index!A1", TextToDisplay:="Back to Index"
        'sht.Range("A1").Hyperlinks(1).Follow NewWindow:=False, AddHistory:=True
NexSht:
    Next sht
    
    'Process Completed
    MsgBox "Index Created"
    
End Sub

