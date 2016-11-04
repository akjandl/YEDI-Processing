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
    Dim dataDir             As String
    Dim scoreFile           As String
    Dim templateFile        As String
    Dim xlsFiles            As Collection
    Dim templateSheets      As Collection
    Dim Locations           As New Dictionary
    Dim curFile             As Variant
    Dim participantID       As String
    
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
    ' Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Open the template files
    Set TemplateBook = Workbooks.Open(templateFile)
    Set ScoreBook = Workbooks.Open(scoreFile)
    
    ' Rename template worksheets & store updated names
    Set templateSheets = HelperFunctions.RenameSheets(TemplateBook, _
        Array("Stroop", "", "", ""), _
        "_updated")
    
    ' Generate list of xls files
    Set xlsFiles = xlsFinder.xlsFinder(dataDir)
    
    ' Populate metadata dictionary with data locations
    Set Locations = metaData.GenerateDictionary()
    
    ' Loop over files
    For Each curFile In xlsFiles
    
        ' Get & store participant ID;
        participantID = HelperFunctions.ExtractID(curFile)
    
        ' Open workbook
        Set CurrentBook = Workbooks.Open(curFile)
        
        ' Copy new template worksheets;
        Call HelperFunctions.CopyWorksheet(TemplateBook, templateSheets, "BRIEF-Parent")
        
        ' TODO: copy data for each _updated sheet
        ' Populate data columns of new worksheets w/data from old;
        ' HelperFunctions.CopyData(SrcRng, DestRng)
        
        ' Check whether final calculated value has an error
        ' Find row in tracker for participant;
        ' HelperFunctions.FindParticipant(ws, ID)
            
            ' If so, log pertinent info in log file
            ' TODO: HelperFunctions.LogError(params?)
            
        ' Compare final calculated values against tracking file
        
            ' If different, overwrite;
            ' HelperFunctions.VerifyAndOverwrite(SrcRng, DestRng)
            
        ' Save & close data file
        
    Next curFile
    
    ' Close Workbooks
    TemplateBook.Close False
    ScoreBook.Close True

    ' Turn alerts back on
    Application.DisplayAlerts = True

End Sub

