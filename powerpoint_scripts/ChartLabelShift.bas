Sub ChartLabelShift()
	'copy chart label style & format from a past data point to latest data point
	'previous data point should be manually edited below
	'script runs on single selection
	
    Dim shp As Shape
    Dim ser As Series
    Dim i As Integer, p As Integer
    Dim pmax As Integer
    Dim poffset as Integer
    
	Let poffset = 3  'Update: number of data points back to copy the label formatting from
    
	Set shp = ActiveWindow.Selection.ShapeRange(1)
    For i = 1 To shp.Chart.SeriesCollection.Count
    
        Set ser = shp.Chart.SeriesCollection(i)
        Let pmax = ser.Points.Count

        ser.Points(pmax).HasDataLabel = True
        With ser.Points(pmax)
            .ApplyDataLabels Type:=xlValue
            .DataLabel.Font.Color = ser.Points(poffset).DataLabel.Font.Color
            .DataLabel.Font.Name = ser.Points(poffset).DataLabel.Font.Name
            .DataLabel.Font.FontStyle = ser.Points(poffset).DataLabel.Font.FontStyle
            .DataLabel.Font.Size = ser.Points(poffset).DataLabel.Font.Size
        End With
        ser.Points(poffset).HasDataLabel = False
    Next i
End Sub