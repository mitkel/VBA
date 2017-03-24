Public Sub AppendFromTextFile(ByRef sEnd As String, _
 ByVal sPath As String, ByVal sName As String, _
 Optional ByVal sSep As String = "")

    Dim myFile As String
    Dim sLine As String
    
    myFile = sPath & sName
    Open myFile For Input As #1
    
    Do Until EOF(1)
        Line Input #1, sLine
        sEnd = sEnd & sLine & sSep
    Loop
    
    Close #1
End Sub

