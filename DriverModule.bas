Attribute VB_Name = "DriverModule"
' Requirements:
'   - Microsoft Scripting Runtime
'   - Microsoft VBScript Regular Expressions 5.5
'
' Worksheets to update:
' number-letter, stroop, stop signal, category switch

Sub Driver()

    Dim TemplateBook        As Workbook
    Dim CurrentBook         As Workbook
    Dim ScoreBook           As Workbook
    Dim LogBook             As Workbook
    Dim ErrorLog            As Worksheet
    Dim curLogRow           As Integer
    Dim errorCount          As Integer
    Dim dataDir             As String
    Dim scoreFile           As String
    Dim templateFile        As String
    Dim suffix              As String
    Dim originalSheets()    As String
    Dim participantID       As String
    Dim dataSetID           As String
    Dim xlsFiles            As Collection
    Dim templateSheets      As Collection
    Dim Locations           As New Dictionary
    Dim curFile             As Variant
    Dim sheetName           As Variant
    Dim i                   As Variant
    Dim valLength           As Integer
    Dim scoreRow            As Range
    Dim fso                 As New FileSystemObject
    
    ' Prompt user to select data root directory, template file, compiled scores
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Select data root directory"
        .Show
        dataDir = .SelectedItems(1)
    End With
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Select template file with correct formulas"
        .Show
        templateFile = .SelectedItems(1)
    End With
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Select the file with the combined compiled scores to verify against"
        .Show
        scoreFile = .SelectedItems(1)
    End With
    
    ' Performance considerations
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Open the template files
    Set TemplateBook = Workbooks.Open(templateFile)
    Set ScoreBook = Workbooks.Open(scoreFile)

    ' Drop error worksheet in macro workbook if exists and recreate & assign
    If ThisWorkbook.Worksheets.Count > 1 Then
        ThisWorkbook.Worksheets(2).Delete
    End If
    
    With ThisWorkbook
        .Worksheets.Add After:=.Worksheets(1)
        .Worksheets(2).name = "Error Log"
    End With
    
    Set ErrorLog = ThisWorkbook.Worksheets("Error Log")
    With ErrorLog
        .Cells(1, 1).Value = "Participant"
        .Cells(1, 2).Value = "Sheet Name"
        .Cells(1, 3).Value = "Number of Errors"
    End With
    curLogRow = 2
    
    ' Rename template worksheets & store updated names
    originalSheets = Split("Stroop,Stop Signal (SSRT Hannah),Category Switch,Number-Letter", ",")
    suffix = "_u"
    Set templateSheets = HelperFunctions.RenameSheets(TemplateBook, originalSheets, suffix)
    
    ' Copy ScoreBook worksheets (2nd Timepoint & Community) & rename w/first word
    Call HelperFunctions.RenameScoreSheets(ScoreBook)
    
    ' Generate list of xls files
    Set xlsFiles = xlsFinder.xlsFinder(dataDir)
    
    ' Populate metadata dictionary with data locations
    Set Locations = metaData.GenerateDictionary()
    
    ' Loop over files
    For Each curFile In xlsFiles
    
        ' Get & store participant ID;
        participantID = HelperFunctions.ExtractID(fso.GetFileName(curFile))
            
        ' Open workbook
        Set CurrentBook = Workbooks.Open(curFile)
        
        ' Copy new template worksheets;
        Call HelperFunctions.CopyWorksheet(TemplateBook, templateSheets, _
            BeforeWs:=CurrentBook.Worksheets("Sentence Completion"))
            
        For Each sheetName In CurrentBook.Worksheets
        
            sheetName.name = Trim(sheetName.name)
            
        Next sheetName
        
        ' Populate data columns of new worksheets w/data from old;
        For Each sheetName In originalSheets
            Call HelperFunctions.CopyData( _
                CurrentBook.Worksheets(sheetName).Range(Locations(sheetName)("Start"), Locations(sheetName)("End")), _
                CurrentBook.Worksheets(sheetName & suffix).Range(Locations(sheetName)("Start"), Locations(sheetName)("End")) _
            )
        Next sheetName
        
        ' Figure out which dataset the participant belongs to (e.g., Community vs 2nd Time Point)
        dataSetID = Split(fso.GetBaseName(fso.GetParentFolderName(fso.GetParentFolderName(curFile))))(0)
        
        ' Check whether final calculated value has an error
        ' Find row in tracker for participant;
        ' HelperFunctions.FindParticipant(ws, ID)
        Set scoreRow = HelperFunctions.FindParticipant(ScoreBook.Worksheets(dataSetID), participantID)
            
        For Each sheetName In CurrentBook.Worksheets
        
            ' Compare final calculated values against tracking file
            If Locations.Exists(sheetName.name) Then
            
                errorCount = 0
            
                For i = 0 To UBound(Locations(sheetName.name)("UserVal"))
                
                    On Error Resume Next:
                    
                    ' If different, overwrite;
                    Call HelperFunctions.VerifyAndOverwrite( _
                        CurrentBook.Worksheets(sheetName.name).Range( _
                            Locations(sheetName.name)("UserVal")(i)), _
                        scoreRow.Cells(Locations(sheetName.name)("CompiledVal")(i)) _
                    )
                    
                    If Err.Number <> 0 Then
                        
                        errorCount = errorCount + 1
                        Err.Clear
                    
                    End If
                    
                Next i
                
                If errorCount > 0 Then
                    
                    ' Log info about participant file, sheet, and error count
                    With ErrorLog
                        With .Cells(curLogRow, 1)
                            .NumberFormat = "@"
                            .Value = participantID
                        End With
                        .Cells(curLogRow, 2).Value = sheetName.name
                        .Cells(curLogRow, 3).Value = errorCount
                    End With
                    curLogRow = curLogRow + 1
                    
                End If
                
            End If
        
        Next sheetName
            
        ' Save & close data file
        ' TODO: change SaveChanges:=True when done debugging
        CurrentBook.Close SaveChanges:=False
        
    Next curFile
    
    ' Close Workbooks
    TemplateBook.Close SaveChanges:=False
    ScoreBook.SaveAs fso.BuildPath(fso.GetFolder(ThisWorkbook.Path), "Updated Compiled Scores")
    ScoreBook.Close SaveChanges:=False

    ' Turn alerts back on
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    ' Autofit column sizing
    ErrorLog.Range(Columns(1), Columns(3)).EntireColumn.AutoFit
    
End Sub

