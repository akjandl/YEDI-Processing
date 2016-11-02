Attribute VB_Name = "DriverModule"
' Dictionary Usage Example
' Requires: Microsoft Scripting Runtime
'
' Dim TestDict            As New Dictionary
' TestDict.Add "Stroop", 8
' Debug.Print TestDict("Stroop")
'
' Idea: Store cell locations, other metadat for verification in Dictionary object

Sub Driver()

    Dim TemplateBook        As Workbook
    Dim CurrentBook         As Workbook
    Dim xlsFiles            As Collection
    Dim curFile             As Variant
    Dim StrFile             As String
    Dim new_file            As String
    Dim Directory           As String
    Dim fso                 As New Scripting.FileSystemObject
        
    ' Performance considerations
    ' Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Use User-input directory else macro workbook's current directory
    Directory = ActiveSheet.Range("D3").Value & "\"
    If Len(Directory) = 0 Then
        Directory = ThisWorkbook.Path & "\"
    End If
    
    ' Open the template files
    Set TemplateBook = Workbooks.Open(Directory & "\" & "EF_scoringtemplate_CORRECTED.xls")
    
    ' Generate list of xls files
    Set xlsFiles = xlsFinder.xlsFinder(Directory & "EF R21")
    
    ' Loop over files
    
        ' Get & store participant ID;
        ' HelperFunctions.ExtractID(FileName)
    
        ' Rename template worksheets
        ' HelperFunctions.RenameSheets(wb, WsNames, Suffix)
        
        ' Copy new template worksheets;
        ' HelperFunctions.CopyWorksheet(SrcWb, SrcWsNames, AfterWs)
        
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
        
        ' Iterate loop

    ' Turn alerts back on
    Application.DisplayAlerts = True

End Sub

