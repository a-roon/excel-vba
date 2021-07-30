Attribute VB_Name = "CombineSheets"
Option Explicit

Sub CombineMultipleSheets()

'Created by Arun Kushvaha
'Last modified 13-Jul-2017

Call A_UsingVBA
'Call CleanData
'Call CreatePivot
MsgBox "Completed!"

End Sub

Private Sub A_UsingVBA()
    Dim myPath As String, FilesInPath As String
    Dim MyFiles() As String
    Dim SourceRcount As Long, Fnum As Long
    Dim mybook As Workbook, BaseWks As Worksheet
    Dim SourceRange As Range, DestRange As Range
    Dim Rnum As Long, CalcMode As Long
    Dim LRnum As Long
    Dim diaFolder As FileDialog
    Dim YesOrNo As String
    
    Application.ScreenUpdating = False
CompileDPT:
    Set diaFolder = Application.FileDialog(msoFileDialogFolderPicker)
    diaFolder.AllowMultiSelect = False
    diaFolder.Title = "Select Source Folder"
    diaFolder.Show

    If diaFolder.SelectedItems.Count = 0 Then
    YesOrNo = MsgBox("Nothing Selected, Are You Sure to Exit?", _
                    vbYesNo, "Folder Not Selected")
    If YesOrNo = vbYes Then GoTo NextCode
    If YesOrNo = vbNo Then GoTo CompileDPT
    End If

    myPath = diaFolder.SelectedItems(1) & "\"
    
    FilesInPath = Dir(myPath & "*.xl*")
    If FilesInPath = "" Then
        MsgBox "No files found"
        Exit Sub
    End If
    
    Fnum = 0
    Do While FilesInPath <> ""
        Fnum = Fnum + 1
        ReDim Preserve MyFiles(1 To Fnum)
        MyFiles(Fnum) = FilesInPath
        FilesInPath = Dir()
    Loop

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With

    Set BaseWks = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    'Set BaseWks = ActiveWorkbook.Worksheets(1)
    BaseWks.Name = "QDR_KPI6"
    Rnum = 1
    
    If Fnum > 0 Then
        For Fnum = LBound(MyFiles) To UBound(MyFiles)
            Set mybook = Nothing
            On Error Resume Next
            Set mybook = Workbooks.Open(myPath & MyFiles(Fnum))
            On Error GoTo 0
            
            If Not mybook Is Nothing Then
                
                On Error Resume Next
                
                If Fnum = 1 Then
                With mybook.Worksheets(1)
                    LRnum = LastRow(mybook.Worksheets(1))
                    Set SourceRange = .Range("A1:J1" & LRnum) 'set the range here
                End With
                Else
                With mybook.Worksheets(1)
                    LRnum = LastRow(mybook.Worksheets(1))
                    Set SourceRange = .Range("A1:J1" & LRnum) 'set the range here
                End With
                End If
                
                If Err.Number > 0 Then
                    Err.Clear
                    Set SourceRange = Nothing
                Else
                    'if SourceRange use all columns then skip this file
                    If SourceRange.Columns.Count >= BaseWks.Columns.Count Then
                        Set SourceRange = Nothing
                    End If
                End If
                On Error GoTo 0

                If Not SourceRange Is Nothing Then

                    SourceRcount = SourceRange.Rows.Count

                    If Rnum + SourceRcount >= BaseWks.Rows.Count Then
                        MsgBox "Sorry there are not enough rows in the sheet"
                        BaseWks.Columns.AutoFit
                        mybook.Close savechanges:=False
                        GoTo ExitTheSub
                    Else

                        'Copy the file name in column K
                        With SourceRange
                        BaseWks.Cells(Rnum, "K"). _
                                    Resize(.Rows.Count).Value = MyFiles(Fnum)
                        End With

                        Set DestRange = BaseWks.Range("A" & Rnum)
                        With SourceRange
                            Set DestRange = DestRange.Resize(.Rows.Count, .Columns.Count)
                        End With
                        'DestRange.Value = SourceRange.Value
                        'SourceRange.Copy DestRange
                        DestRange.Value = SourceRange.Value
                        Rnum = Rnum + SourceRcount
                    End If
                End If
                mybook.Close savechanges:=False
            End If

        Next Fnum
        BaseWks.UsedRange.WrapText = False
        BaseWks.Columns.AutoFit
        BaseWks.Cells(1, "K").Value = "File Name" 'Set the filename column
    End If
NextCode:
    Set diaFolder = Nothing
    
ExitTheSub:
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
    Application.ScreenUpdating = True
End Sub

