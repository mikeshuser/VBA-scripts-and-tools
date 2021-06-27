Sub FormatLineChart()
	'automatic formatting of powerpoint line chart
	'need to specify the chart point numbers to be formatted
	'example: points_ndx = (1, 13) -> to format most recent month and same month last year 

    Dim shp As Shape
    Dim ser As Series
    Dim i As Integer
	Dim pnt as Integer
	Dim points_ndx As Variant
    
	Let points_ndx = Array(1, 13, 25)
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    
    For i = 1 To shp.Chart.SeriesCollection.Count
    
        Set ser = shp.Chart.SeriesCollection(i)
        ser.DataLabels.ShowValue = False
        ser.MarkerStyle = xlMarkerStyleNone
        ser.Format.ThreeD.BevelTopDepth = 0
        ser.Format.ThreeD.BevelTopInset = 0
        ser.Format.Fill.ForeColor = ser.Format.Line.ForeColor
		ser.Format.Shadow.Visible = msoFalse
        
		For each pnt in points_ndx
			With ser.Points(pnt)
				.MarkerStyle = xlMarkerStyleCircle
				.MarkerSize = 4
				.ApplyDataLabels Type:=xlValue
				.DataLabel.Font.Color = RGB(89, 89, 89)
				.DataLabel.Font.Name = "Arial"
				.DataLabel.Font.FontStyle = "Bold"
				.DataLabel.Font.Size = 8
				.Format.Line.ForeColor = ser.Format.Line.ForeColor
				.Format.Shadow.Visible = msoFalse
			End With
		Next pnt     
    Next i
End Sub