Sub FormatTable()
	'format data table values. Need to select table before running
    Dim shp As Shape
    Dim row As Integer, col As Integer
    Dim val As String
    
    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If shp.HasTable Then
        For row = 1 To shp.Table.Rows.Count
            For col = 1 To shp.Table.Columns.Count
                Let val = shp.Table.Cell(row, col).Shape.TextFrame.TextRange.Text
                If IsNumeric(val) = True Then
                    With shp.Table.Cell(row, col).Shape
                        .TextFrame.TextRange.Font.Name = "Arial"
                        .TextFrame.TextRange.Font.Size = "10"
                        .TextFrame.TextRange.Font.Bold = True
                        .TextFrame.HorizontalAnchor = msoAnchorCenter
                        .TextFrame.VerticalAnchor = msoAnchorMiddle
                    
						If val >= 0.015 Then
							.TextFrame.TextRange.Font.Color = RGB(118, 192, 67)
						ElseIf val <= -0.015 Then
							.TextFrame.TextRange.Font.Color = RGB(255, 0, 0)
						Else
							.TextFrame.TextRange.Font.Color = RGB(127, 127, 127)
						End If
                    
						.TextFrame.TextRange.Text = Format(val, "+#%;-#%;0%")
					End With
                End If
            Next col
        Next row
    End If
End Sub