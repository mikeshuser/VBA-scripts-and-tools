Sub FormatGapChart()
	'applies standard gap bar chart colouring (green>=5%, red<=-5%, -5%<grey<5%)
	'runs on all selected charts
    Dim shp As Shape
    Dim ser As Series
    Dim p As Integer, i As Integer
    
    For Each shp In ActiveWindow.Selection.ShapeRange
		For i = 1 To shp.Chart.SeriesCollection.Count
			Set ser = shp.Chart.SeriesCollection(i)
			For p = 1 To ser.Points.Count
				If ser.Values(p) >= 0.045 Then 
					ser.Points(p).Interior.Color = RGB(118, 192, 67)
				
				ElseIf ser.Values(p) <= -0.045 Then 
					ser.Points(p).Interior.Color = RGB(255, 0, 0)
				
				Else 
					ser.Points(p).Interior.Color = RGB(127, 127, 127)
				End If
			Next p
		Next i
    Next shp  
End Sub