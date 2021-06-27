Function LetterToNum(ByVal strSource As String) As String
	'return the alphabet position of a letter. Useful for certain indexing applications
	'ex: LetterToNum("a") = 1, LetterToNum("b") = 2, etc...
	
    Dim i As Integer
    Dim strResult As String
    
    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
            Case 65 To 90:
                strResult = strResult & Asc(Mid(strSource, i, 1)) - 64
            Case Else
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    LetterToNum = strResult

End Function