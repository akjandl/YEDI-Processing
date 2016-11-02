Attribute VB_Name = "Conversion"
Sub Convert_XLS_XLSX()

    Dim oBook       As Workbook
    Dim StrFile     As String
    Dim macro_name  As String
    Dim new_file    As String
    Dim Directory   As String
    Dim fso         As New Scripting.FileSystemObject
        
    ' Performance considerations
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    ' There was a semicolon here
    ' Use User-input directory else macro workbook's current directory
    Directory = ActiveSheet.Range("D3").Value & "\"
    If Len(Directory) = 0 Then
        Directory = ThisWorkbook.Path & "\"
    End If
    
    ' Other vars
    macro_name = ThisWorkbook.Name
    StrFile = Dir(Directory & "*.xls")
    
    ' Loop over xls files
    Do While Len(StrFile) > 0
        
        ' change this to check that is actually *.xls
        If StrFile <> macro_name Then
            
            Debug.Print "Converting " & StrFile & "..."
            new_file = Directory & fso.GetBaseName(StrFile)
            Workbooks.Open Directory & "\" & StrFile
            Workbooks(StrFile).SaveAs new_file & ".xlsx", XlFileFormat.xlOpenXMLWorkbook
            ActiveWorkbook.Close False
        
        End If
        
        StrFile = Dir
        
    Loop
    
    ' Turn alerts back on
    Application.DisplayAlerts = True

End Sub
