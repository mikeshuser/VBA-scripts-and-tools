Sub SortByLast()
	'sort the selected table by the last column in descending order
	'assumes header row is not included
	
	Dim rng As Range
	Dim keyrng As Range
	Dim rowcount As Integer, colcount As Integer

	Set rng = Selection
	Let rowcount = rng.rows.Count
	Let colcount = rng.Columns.Count
	Set keyrng = Range(Selection.Cells(1, colcount), Selection.Cells(rowcount, colcount))

	rng.Sort key1:=keyrng, order1:=xlDescending, Header:=xlNo
End Sub


Sub SortBySecond()
	'sort the selected table by the second column in descending order
	'assumes header row is not included
	
	Dim rng As Range
	Dim keyrng As Range
	Dim rowcount As Integer

	Set rng = Selection
	Let rowcount = rng.rows.Count
	Set keyrng = Range(Selection.Cells(1, 2), Selection.Cells(rowcount, 2))

	rng.Sort key1:=keyrng, order1:=xlDescending, Header:=xlNo
End Sub