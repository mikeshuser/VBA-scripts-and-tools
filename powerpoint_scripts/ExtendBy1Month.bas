Sub ExtendBy1Month()
	'conveniance macro to add an additional month to all selected charts
	'works if data is vertical or horizontal orientation
	'currently need to specify month value to append to data range
	'TODO: add automatic date parsing to eliminate manual month update
	
    Dim chData As ChartData
    Dim dataSheet As Worksheet
    Dim shp As Shape
    Dim sRow As Integer, eRow As Integer
    Dim sCol As Integer, eCol As Integer
    Dim rowCount As Integer, colCount As Integer
    Dim month As String
    
    Let month = "Dec '20" 'update month manually for now
    
    For Each shp In ActiveWindow.Selection.ShapeRange
        If shp.HasChart Then
            If shp.Chart.ChartType = xlLine Then
                Set chData = shp.Chart.ChartData
                chData.Activate
                Set dataSheet = chData.Workbook.Worksheets(1)
                
                Let rowCount = dataSheet.Range("Table1").Rows.Count
                Let colCount = dataSheet.Range("Table1").Columns.Count
        
                If rowCount > colCount Then
                    Let eRow = dataSheet.Range("A1").End(xlDown).row
                    Let sRow = eRow - 12
                
                    dataSheet.Cells(eRow + 1, 1).Value = month
                    dataSheet.Rows(sRow).EntireRow.Hidden = True
                Else
                    Let eCol = dataSheet.Range("A1").End(xlToRight).Column
                    Let sCol = eCol - 12
    
                    dataSheet.Cells(1, eCol + 1).Value = month
                    dataSheet.Columns(sCol).EntireColumn.Hidden = True
                End If
                chData.Workbook.Close
            End If
        End If
    Next shp
End Sub

