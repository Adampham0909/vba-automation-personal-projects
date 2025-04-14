Attribute VB_Name = "Module1"
Sub FilteredAndSummarizeSales()
    Dim wsFiltered As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, ouputRow As Long
    Dim totalValue As Double
    
    
    On Error Resume Next
    Set wsFiltered = ThisWorkbook.Sheets("filtered_sales")
    If wsFiltered Is Nothing Then
        Set wsFiltered = ThisWorkbook.Sheets.Add
        wsFiltered.Name = "filtered_sales"
    Else
        wsFiltered.Cells.ClearContents
    End If
    On Error GoTo 0
    
    wsFiltered.Range("A1:F1").Value = Array("Date", "Product", "Quantity", "Unit Price", "Total", "Sheet Source")
    outputRow = 2
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "filtered_sales" Then
            lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            For i = 2 To lastRow
                totalValue = ws.Cells(i, 5).Value
                If totalValue >= 100 Then
                    ws.Range("A" & i & ":E" & i).Copy Destination:=wsFiltered.Range("A" & outputRow)
                    wsFiltered.Cells(outputRow, 6).Value = ws.Name
                    outputRow = outputRow + 1
                End If
            Next i
            
        End If
    Next ws
        
End Sub
