Attribute VB_Name = "DriverModule"
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

    ' Turn alerts back on
    Application.DisplayAlerts = True

End Sub

