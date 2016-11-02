Attribute VB_Name = "HelperFunctions"
Function ExtractID(FileName As String) As String

    Dim Regex           As New VBScript_RegExp_55.RegExp
    Dim Matches         As Object
    Dim UserID          As String
    
    With Regex
        .IgnoreCase = True
        .Pattern = "^d?t?c?\s?([\d]{1,4}).*$"
        If .test(FileName) Then
            Set Matches = .Execute(FileName)
            UserID = Matches(0).SubMatches(0)
        Else
            MsgBox FileName & " does not contain a UserID"
        End If
    End With
    
    Set Regex = Nothing  ' free memory?
    ExtractID = UserID
    
End Function
