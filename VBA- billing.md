**insert and color ( alt + f8 then code then f5)**

Sub InsertGrayRowsWhenDateChanges()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    ' Loop from bottom to top, only from row 7 downwards
    For i = lastRow To 7 Step -1
        ' Compare date values in col D
        If ws.Cells(i, "D").Value <> ws.Cells(i - 1, "D").Value Then
            ' Insert a new row
            ws.Rows(i).Insert Shift:=xlDown
            ' Fill columns A:M with gray
            ws.Range("A" & i & ":M" & i).Interior.Color = RGB(217, 217, 217)
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub
