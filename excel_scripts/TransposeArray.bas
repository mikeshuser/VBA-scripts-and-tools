Option Base 0

Function TransposeArray(arr As Variant) As Variant
	'Transposes array. Assumes option base 0
	
	Dim X As Long, Y as Long
	Dim Xbound As Long, Ybound as Long
	Dim tmpArray As Variant
	
    Let Xbound = UBound(arr, 1)
    Let Ybound = UBound(arr, 0)
    ReDim tmpArray(Xbound, Ybound) 
    For X = 0 To Xbound
        For Y = 0 To Ybound 
            tmpArray(X, Y) = arr(Y, X) 
        Next Y 
    Next X 
    TransposeArray = tmpArray 
End Function