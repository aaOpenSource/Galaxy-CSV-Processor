VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Galaxy CSV Processor"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12735
   OleObjectBlob   =   "frmMain.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Notes()

' For the code to run you must add a reference to ADO, Microsoft Active Data Objects.  I typically run with 2.6 but you can try lower.
'
'
'
'
'
'
'
'
'
'
'
'


End Sub

Private Sub btnClearLog_Click()
    ClearLog
End Sub

Private Sub ClearLog()
    txtLog.Text = ""
End Sub

Private Sub Log(Text As String)
    txtLog.Text = txtLog.Text & vbCrLf & Text
End Sub

Private Sub btnImport_Click()
    ClearLog
    ImportGalaxyCSV (Application.GetOpenFilename("Comma Separated Values (*.csv),*.csv"))
End Sub

Sub ImportGalaxyCSV(OpenFileName As String)


    If OpenFileName = "False" Then
        Log ("No valid file selected")
    Else
    
        Log ("Opening File " & OpenFileName)
            
        Workbooks.OpenText Filename:= _
            OpenFileName, _
            Origin:=65001, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
            xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False _
            , Comma:=True, Space:=False, Other:=False, TrailingMinusNumbers:=True
        
        ' Seperate the Templates to individual worksheets
        Log ("Separating templates")
        
        SeparateGalaxyDump
        
        Log ("Saving as Excel File")
        
        Log SaveFileName(OpenFileName)
        
        Application.ActiveWorkbook.SaveAs Filename:=SaveFileName(OpenFileName), FileFormat:=xlWorkbookDefault
        
        Log ("Saved as Excel File")
        
    End If

End Sub

Function SaveFileName(InFileName As String) As String

    ' Calculate the correct file name to save as
    
    ' Find the trailing .
    Dim TrailingDotLoc As Integer
    
    TrailingDotLoc = InStrRev(InFileName, ".")
    
    If TrailingDotLoc >= 0 Then
        SaveFileName = Left(InFileName, TrailingDotLoc - 1) & ".XLSX"
    Else
        SaveFileName = InFileName & ".XLSX"
    End If
            
    
End Function

Sub SeparateGalaxyDump()

    Dim TemplateCells As Collection
    Dim TemplateCellAddress As Variant
    Dim TemplateCell As Range
    Dim SelectRange As Range
    Dim WorkingSheet As Worksheet
    Dim NewWorksheet As Worksheet
    
    Application.ScreenUpdating = False
    
    
    Set WorkingSheet = Worksheets(GetActiveSheet)
        
    ' Get the template cells using the find method
    Set TemplateCells = GetTemplateCellAddesses(WorkingSheet)

    ' Loop through them and push the contiguous region to new worksheets
    For Each TemplateCellAddress In TemplateCells
        
        ' Get the template cell
        Set TemplateCell = WorkingSheet.Range(TemplateCellAddress)
        
        Log ("Creating " & TemplateCell.Text)
        
        ' Create a new worksheet based on
        Set NewWorksheet = MoveTemplateToNewWorksheet(TemplateCell)
        
    Next
    
    ' Delete the working sheet when done
    Application.DisplayAlerts = False
    WorkingSheet.Delete
    Application.DisplayAlerts = True
    
    ' Activate the first worksheet
    Application.Worksheets(1).Activate
    
    Application.ScreenUpdating = True
    
End Sub

Function GetActiveSheet() As String
  GetActiveSheet = ActiveSheet.Name
End Function

Sub DeleteWorksheetIfExists(WorkSheetName As String)

    Dim ws As Worksheet
    Dim wsName As String
    
    
    wsName = UCase(WorkSheetName)
    
    For Each ws In ActiveWorkbook.Sheets
        If UCase(ws.Name) = wsName Then
            Application.DisplayAlerts = False
            'Debug.Print "Deleting " & WorksheetName
            ws.Delete
            Application.DisplayAlerts = True
            Exit For
        End If
    Next
    
End Sub


Function MoveTemplateToNewWorksheet(TemplateCell As Range) As Worksheet

    Dim NewWorksheet As Worksheet
    Dim TemplateName As String
       
    TemplateName = GetTemplateName(TemplateCell.Text)
    
    ' Delete the old one if it existss
    DeleteWorksheetIfExists (TemplateName)
    
    Set NewWorksheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    
    NewWorksheet.Name = TemplateName
    
    ' Copy the contents of the template region to the new worksheet
    TemplateCell.CurrentRegion.Copy NewWorksheet.Cells(1, 1)
    
    If ckAutoWiden.Value Then
       NewWorksheet.Cells.EntireColumn.AutoFit
    End If
    
        
    Set MoveTemplateToNewWorksheet = NewWorksheet
    
End Function

Function GetTemplateName(InValue As String) As String
    GetTemplateName = Right(InValue, Len(InValue) - Len(":Template="))
End Function

Function GetTemplateCellAddesses(InWorksheet As Worksheet) As Collection

    Dim SearchCol As Range
    Dim FoundCell As Range
    Dim FirstAddress
    
    Dim FoundTemplateCells As Collection
    Set FoundTemplateCells = New Collection
    
    Set SearchCol = InWorksheet.Range("A:A")
    
    With SearchCol
        Set FoundCell = .Cells.Find(What:=":Template=")
        If Not FoundCell Is Nothing Then
            FirstAddress = FoundCell.Address
            Do
                FoundTemplateCells.Add (FoundCell.Address)
                Set FoundCell = .Cells.FindNext(FoundCell)
            Loop Until FoundCell Is Nothing Or FoundCell.Address = FirstAddress
        End If
    End With
    
    Set GetTemplateCellAddesses = FoundTemplateCells

End Function

Function CreateExportWorksheet(WorkSheetName) As Worksheet

    Dim NewWorksheet As Worksheet
    
    ' Delete the old one if it existss
    DeleteWorksheetIfExists (WorkSheetName)
    
    Set NewWorksheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    
    NewWorksheet.Name = WorkSheetName
            
    Set CreateExportWorksheet = NewWorksheet
    
End Function



Function GetLoggedOnUser() As String

    
    Dim strLen As Long
    Dim strtmp As String * 256
    Dim strUserName As String

    strLen = 255
    GetUserName strtmp, strLen
    
    strUserName = Left(strtmp, strLen)
    
    'strUserName = Trim$(TrimNull(strtmp))
    GetLoggedOnUser = strUserName
    
End Function

Function CopyTemplatesToExportWorksheet(ByRef ExportWorksheet) As Boolean

    Dim ws As Worksheet
    Dim FirstBlankRow As Integer
    Dim PastedRowCount As Integer
    Dim ListCount As Integer
    Dim Index As Integer
    
    Dim SheetsList As Collection
    Set SheetsList = New Collection
        
    ' Determine the list of sheets to export
    
    If optAllSheets.Value Then
        For Each ws In ActiveWorkbook.Sheets
            If (GetTemplateCellAddesses(ws).Count > 0) And (ws.Name <> ExportWorksheet.Name) Then
                SheetsList.Add ws
            End If
        Next
    End If
    
    If optSelectedSheets.Value Then
        
        ListCount = lstSheets.ListCount
        
        For Index = 0 To ListCount - 1
        
            If lstSheets.Selected(Index) Then
                SheetsList.Add ActiveWorkbook.Sheets(lstSheets.List(Index))
            End If
            
        Next
        
    End If
    
    ' If none to export then bail
    If SheetsList.Count = 0 Then
        Log "No sheets selected for export."
        CopyTemplatesToExportWorksheet = False
        Exit Function
    End If
                
    ' Add a note at the top about who and when
    ExportWorksheet.Range("A1").Value = ";Export prepared on " & DateTime.Now & " by " & GetLoggedOnUser()
    
    FirstBlankRow = 3
                
    ' Otherwise loop through the collection and export the sheets
    For Each ws In SheetsList
        If (ws.Name <> ExportWorksheet.Name) Then
            Log ("Copying " & ws.Name & " for export")
            ws.Range("A1").CurrentRegion.Copy ExportWorksheet.Cells(FirstBlankRow, 1)
            PastedRowCount = ExportWorksheet.Range("A" & FirstBlankRow).CurrentRegion.Rows.Count
            FirstBlankRow = PastedRowCount + FirstBlankRow + 1
        End If
    Next
    
    CopyTemplatesToExportWorksheet = True
    
End Function


Private Sub btnExport_Click()
    
    ClearLog
    
    Dim LocalExportWorksheet As Worksheet
    Dim ExportWorkSheetName As String
    
    ExportWorkSheetName = "TempExport"
    
    Log ("Creating Export Worksheet" & ExportWorkSheetName)
    Set LocalExportWorksheet = CreateExportWorksheet(ExportWorkSheetName)
    
    Log ("Copying Worksheets to " & ExportWorkSheetName)
    If Not CopyTemplatesToExportWorksheet(LocalExportWorksheet) Then
        Log ("Error copying sheets.")
        Exit Sub
    End If
        
    Log ("Exporting " & ExportWorkSheetName)
    ExportSpecificWorksheet LocalExportWorksheet
    
    Log ("Deleting Temporary Worksheet " & ExportWorkSheetName)
    DeleteWorksheetIfExists (LocalExportWorksheet.Name)
    
    Log ("Complete")
    
End Sub

Function ExportSpecificWorksheet(ByRef ExportWorksheet)

    Dim oldname As String
    Dim oldpath As String
    Dim oldFormat As Integer
    Dim ExportFileName As String
    
    ExportFileName = Application.GetSaveAsFilename(fileFilter:="CSV File (*.csv), *.csv")
      
    ' Turn off alerts
    Application.DisplayAlerts = False  'avoid safetey alert
    
  ' Get the old name, path, and format, then save as CSV, then save back.
    With ActiveWorkbook
      oldname = .Name
      oldpath = .Path
      oldFormat = .FileFormat
    End With

    ' Export the CSV
    SaveUTF8CSV ExportWorksheet, ExportFileName
     
    Log ("Saved to " & ExportFileName)
    
    ' Save Back as Old Format
    ActiveWorkbook.SaveAs Filename:=oldpath + "\" + oldname, FileFormat:=oldFormat
    
    Application.DisplayAlerts = True
    
End Function

Private Sub SaveUTF8CSV(ByRef ExportWorksheet, Filename As String)

    Dim ColumnCount, RowCount As Integer, columnNumber, rowNumber As Integer
    Dim line, cellText As String
    
    Dim ObjStream As ADODB.Stream
    
    ' init stream
    Set ObjStream = New ADODB.Stream
    ObjStream.Open
    ObjStream.Charset = "utf-8"
    ObjStream.Type = adTypeText
    ObjStream.LineSeparator = adCRLF
    ObjStream.Position = 0
    
    ' Get the worksheet into a CSV form and write to object stream
    GetWorksheetCSVToStream ExportWorksheet, ObjStream
    
    ObjStream.SaveToFile Filename, adSaveCreateOverWrite
    
    ' close up and return
    ObjStream.Close
    Set ObjStream = Nothing

End Sub

Private Function GetWorksheetCSVToStream(ByRef ExportWorksheet, ByRef ObjStream As ADODB.Stream)

    Dim ColumnCount, RowCount As Integer, columnNumber, rowNumber As Integer
    Dim line, cellText As String
    Dim TagRow As Boolean
    
    ColumnCount = ExportWorksheet.UsedRange.Columns.Count
    RowCount = ExportWorksheet.UsedRange.Rows.Count
    
    With ExportWorksheet.UsedRange
     For rowNumber = 1 To RowCount
     
       line = ""
       
       ' Determine if this is a row with a tag
       TagRow = IsTagRow(.Cells(rowNumber, 1).Text)
       
       For columnNumber = 1 To ColumnCount
       
        cellText = .Cells(rowNumber, columnNumber).Text
        
        If TagRow Then
            cellText = """" & Replace(cellText, """", """""") & """"
        End If
        
        ' Add any other special cases to manage here
        
        ' Add the trailing ,
        line = line & cellText & ","
        
       Next columnNumber
      ObjStream.WriteText line, adWriteLine
     Next rowNumber
    End With
    
End Function

Private Function IsTagRow(CellData As String)

' Determine if this row is a row with tag data by inspecting the first cell's data

    If CellData = "" Then
        IsTagRow = False
        Exit Function
    End If
    
    If Left(CellData, 1) = ";" Then
        IsTagRow = False
        Exit Function
    End If
    
    If Left(CellData, 1) = ":" Then
        IsTagRow = False
        Exit Function
    End If
    
    IsTagRow = True
    

End Function

Private Sub ckAutoWiden_Click()

End Sub

Private Sub optAllSheets_Click()

    lstSheets.Visible = (optSelectedSheets)
    
End Sub

Private Sub optSelectedSheets_Click()

    lstSheets.Visible = (optSelectedSheets)

    Dim ws As Worksheet
    
    lstSheets.Clear

    For Each ws In ActiveWorkbook.Sheets
        If (GetTemplateCellAddesses(ws).Count > 0) Then
            lstSheets.AddItem ws.Name
        End If
    Next
    
End Sub

Private Sub TabStrip1_Change()

End Sub
