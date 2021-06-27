Function StripNonNumeric(val As Variant) As Variant
    'convert phone numbers of varying formats to single unified format
	
    Dim StartString As String
    Dim PhoneNumber As String
    Dim i As Integer
    
    StartString = val
    PhoneNumber = ""
    For i = 1 To Len(StartString)
        If IsNumeric(Mid(StartString, i, 1)) Then
            PhoneNumber = PhoneNumber & Mid(StartString, i, 1)
        End If
    Next
    
    StripNonNumeric = PhoneNumber
End Function