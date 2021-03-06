Attribute VB_Name = "MetaData"
Function GenerateDictionary() As Dictionary

    ' Function to encapsulate all the metadata involving data
    ' fields and data locations for reference in the rest of the
    ' program. Up to 4 keys may be defined:
    '   - Start: start of data range to copy to/from
    '   - End: end of data range to copy to/from
    '   - UserVal: Array of score locations to check against compiled
    '   - CompiledVal: Array of column numbers to locate values
    
    ' Requires: HelperFunctions.AddEntry
    ' Returns: populated Dictionary object
    
    Dim metaData            As New Dictionary
    
    ' Sentence Completion
    Set metaData = HelperFunctions.AddEntry(metaData, "Sentence Completion", "UserVal", Array("O3"))
    Set metaData = HelperFunctions.AddEntry(metaData, "Sentence Completion", "CompiledVal", Array(2))

    ' Stroop
    Set metaData = HelperFunctions.AddEntry(metaData, "Stroop", "Start", "F2")
    Set metaData = HelperFunctions.AddEntry(metaData, "Stroop", "End", "I340")
    Set metaData = HelperFunctions.AddEntry(metaData, "Stroop", "UserVal", Array("V2", "V4"))
    Set metaData = HelperFunctions.AddEntry(metaData, "Stroop", "CompiledVal", Array(3, 4))
     
    ' Antisaccade
    Set metaData = HelperFunctions.AddEntry(metaData, "Antisaccade", "UserVal", Array("F1", "F5"))
    Set metaData = HelperFunctions.AddEntry(metaData, "Antisaccade", "CompiledVal", Array(5, 6))
    
    ' Stop Signal
    Set metaData = HelperFunctions.AddEntry(metaData, "Stop Signal (SSRT Hannah)", "Start", "E3")
    Set metaData = HelperFunctions.AddEntry(metaData, "Stop Signal (SSRT Hannah)", "End", "I307")
    Set metaData = HelperFunctions.AddEntry(metaData, "Stop Signal (SSRT Hannah)", "UserVal", Array("Q2", "Q3"))
    Set metaData = HelperFunctions.AddEntry(metaData, "Stop Signal (SSRT Hannah)", "CompiledVal", Array(7, 8))
    
    ' Category Switch
    Set metaData = HelperFunctions.AddEntry(metaData, "Category Switch", "Start", "D3")
    Set metaData = HelperFunctions.AddEntry(metaData, "Category Switch", "End", "K254")
    Set metaData = HelperFunctions.AddEntry(metaData, "Category Switch", "UserVal", Array("T2", "T3", "T4", "T5"))
    Set metaData = HelperFunctions.AddEntry(metaData, "Category Switch", "CompiledVal", Array(9, 10, 11, 12))
    
    ' Color-Shape
    Set metaData = HelperFunctions.AddEntry(metaData, "Color-Shape", "UserVal", Array("M3", "M4", "M5", "M6"))
    Set metaData = HelperFunctions.AddEntry(metaData, "Color-Shape", "CompiledVal", Array(13, 14, 15, 16))

    ' Number-Letter
    Set metaData = HelperFunctions.AddEntry(metaData, "Number-Letter", "Start", "C3")
    Set metaData = HelperFunctions.AddEntry(metaData, "Number-Letter", "End", "F402")
    Set metaData = HelperFunctions.AddEntry(metaData, "Number-Letter", "UserVal", Array("O3", "O4", "O5", "O6"))
    Set metaData = HelperFunctions.AddEntry(metaData, "Number-Letter", "CompiledVal", Array(17, 18, 19, 20))
    
    ' Keep Track
    Set metaData = HelperFunctions.AddEntry(metaData, "Keep Track", "UserVal", Array("G2"))
    Set metaData = HelperFunctions.AddEntry(metaData, "Keep Track", "CompiledVal", Array(21))
    
    ' Letter Memory
    Set metaData = HelperFunctions.AddEntry(metaData, "Letter Memory", "UserVal", Array("Q1"))
    Set metaData = HelperFunctions.AddEntry(metaData, "Letter Memory", "CompiledVal", Array(22))
    
    ' 2-back
    Set metaData = HelperFunctions.AddEntry(metaData, "2-back", "UserVal", Array("G1"))
    Set metaData = HelperFunctions.AddEntry(metaData, "2-back", "CompiledVal", Array(23))
    
    ' WASI
    Set metaData = HelperFunctions.AddEntry(metaData, "WASI", "UserVal", Array("B35", "B37", "B39"))
    Set metaData = HelperFunctions.AddEntry(metaData, "WASI", "CompiledVal", Array(24, 25, 26))
    
    ' BRIEF-SR
    ' Set metaData = HelperFunctions.AddEntry(metaData, "BRIEF-SR", "UserVal", Array("J3", "J4", "J5", "J6", "J7", "J8", "J9", "J10", "J11", "J12", "J13", "J14", "J15"))
    ' Set metaData = HelperFunctions.AddEntry(metaData, "BRIEF-SR", "CompiledVal", Array(35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47))
    
    ' BRIEF-Parent
    ' Set metaData = HelperFunctions.AddEntry(metaData, "BRIEF-Parent", "UserVal", Array("J3", "J4", "J5", "J6", "J7", "J8", "J9", "J10", "J11", "J12", "J13"))
    ' Set metaData = HelperFunctions.AddEntry(metaData, "BRIEF-Parent", "CompiledVal", Array(48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58))
     
    Set GenerateDictionary = metaData

End Function
