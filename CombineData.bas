Attribute VB_Name = "Module2"
Sub TongHopDuLieuTuNhieuSheet()

    Dim wsTongHop As Worksheet
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim dongBatDau As Long
    Dim dataRange As Range
    Dim cell As Range

    ' Kh?i t?o
    Set wsTongHop = ThisWorkbook.Sheets("tonghop")
    dongBatDau = 2
    ' Xoá d? li?u cu
    wsTongHop.Range("A2:F1000").ClearContents
    
    wsTongHop.Range("A1:F1").Value = Array("Date", "Product", "Quantity", "Unit Price", "Total", "Store #")

    ' Duy?t qua t?ng sheet
    
    For Each ws In ThisWorkbook.Sheets
        sheetName = ws.Name
            If sheetName <> "tonghop" Then
                lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                If lastRow > 1 Then
                    Set dataRange = ws.Range("A2:E" & lastRow)
                    
                    dataRange.Copy wsTongHop.Cells(dongBatDau, "A")
                    For Each cell In wsTongHop.Range("F" & dongBatDau & ":F" & dongBatDau + dataRange.Rows.Count - 1)
                        cell.Value = sheetName
                    Next cell
                    
                    dongBatDau = dongBatDau + dataRange.Rows.Count
                End If
            End If
    Next ws
    
    MsgBox "All data were combined successfully!", vbInformation

End Sub

