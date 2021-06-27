Public Function DictLookup(lookupRange As Range, refRange As Range, dataCol As Long) As Variant
	'search function similar to vlookup, but inserts data into a dictionary before searching
	'this results in a huge speedup when performing many thousands of searches
	'use as array formula

	Dim dict As Object
	Dim tgtRow As Range
	Dim i As Long, j As Long
	Dim vResults() As Variant

	Set dict = CreateObject("Scripting.Dictionary")
	For Each tgtRow In refRange.Columns(1).Cells
		dict.Add tgtRow.Value, tgtRow.Offset(0, dataCol - 1).Value
	Next tgtRow

	ReDim vResults(1 To lookupRange.Rows.Count, 1 To lookupRange.Columns.Count) As Variant
	For i = 1 To lookupRange.Rows.Count
		For j = 1 To lookupRange.Columns.Count
			If dict.Exists(lookupRange.Cells(i, j).Value) Then
				vResults(i, j) = dict(lookupRange.Cells(i, j).Value)
			End If
		Next j
	Next i

	DictLookup = vResults
End Function
