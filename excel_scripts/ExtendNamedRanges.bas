Sub ExtendNamedRanges(row_extend As Boolean)
    'extend all named ranges in workbook by 1 row or 1 column

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim i As Long
    Dim resizedrng As Range
    Dim nr As Name
    
    Set wb = ThisWorkbook
    For i = 1 To wb.Names.Count
        Set nr = wb.Names.Item(i)
        Set ws = nr.RefersToRange.Worksheet

        With ws.Range(nr.Name)
            If row_extend = True Then
                Set resizedrng = .Resize(.Rows.Count + 1, .Columns.Count)
            Else
                Set resizedrng = .Resize(.Rows.Count, .Columns.Count + 1)
            End If
        End With
        nr.RefersTo = resizedrng
    Next i
End Sub