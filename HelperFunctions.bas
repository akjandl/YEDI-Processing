Attribute VB_Name = "HelperFunctions"
Function CopyData(SrcRng As Range, DestRng As Range)

    ' `SrcRng` contains the data that will be copied over to `DestRng`
    ' n.b., the ranges *must* be the same size
    ' Returns: Nothing
    
    SrcRng.Copy DestRng

End Function

Function CopyWorksheet(SrcWb As Workbook, SrcWsNames As Variant, _
    AfterWs As Worksheet)

    ' Takes the source workbook, source worksheet names (to copy)
    ' and a worksheet in the destination workbook after which
    ' to paste the source worksheet(s)
    ' Returns: Nothing
    
    SrcWb.Worksheets(SrcWsNames).Copy After:=AfterWs

End Function

Function ExtractID(FileName As String) As String

    Dim Regex           As New VBScript_RegExp_55.RegExp
    Dim Matches         As Object
    Dim ParticipantID   As String
    
    With Regex
        .IgnoreCase = True
        .Pattern = "^d?t?c?\s?([\d]{1,4}).*$"
        If .Test(FileName) Then
            Set Matches = .Execute(FileName)
            ParticipantID = Matches(0).SubMatches(0)
        Else
            MsgBox FileName & " does not contain a Participant ID"
        End If
    End With
    
    Set Regex = Nothing  ' free memory?
    ExtractID = ParticipantID
    
End Function

Function FindParticipant(ws As Worksheet, ID As String) As Range

    ' Takes a worksheet and an ID string as arguments
    ' Returns: a Range object representing the row matching the ID
    
    Set FindParticipant = ws.Columns(1).Find(ID, LookIn:=xlValues).EntireRow
    
End Function

Function RenameSheets(wb As Workbook, WsNames As Variant, Suffix As String) _
    As Collection
    
    ' Takes a workbook object, an array of worksheet names to change,
    ' and a suffix to append to each worksheet name
    ' Returns: a collection of the updated worksheet names
    
    Dim newNames        As New Collection
    Dim name            As Variant
    Dim newName         As String
    
    For Each name In WsNames
    
        newName = name & Suffix
        newNames.Add (newName)
        wb.Worksheets(name).name = newName
    
    Next name
    
    Set RenameSheets = newNames
    
End Function

Function VerifyAndOverwrite(SrcRng As Range, DestRng As Range)

    ' Takes two range objects, compares their values, and overwrites
    ' TargetRng values with those of SourceRng if they are not equal
    ' Returns: Nothing (makes changes in place)
    
    If SrcRng.Value <> DestRng.Value Then
        
        TargetRng.Value = SourceRng.Value
        
    End If

End Function
