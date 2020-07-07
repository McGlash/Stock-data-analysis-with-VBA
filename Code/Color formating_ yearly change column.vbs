Attribute VB_Name = "Module4"
Sub marker()

'initiating script across worksheets

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    
ws.Activate

    'Loop!
    
    Dim i As Integer
    
        For i = 2 To Range("I2", Range("I1").End(xlDown)).Rows.Count + 1
        
            If Cells(i, 10).Value > 0 Then
        
                Cells(i, 10).Interior.ColorIndex = 4 'positive=green
                
            ElseIf Cells(i, 10).Value = 0 Then
        
                Cells(i, 10).Interior.ColorIndex = 15 'neutral=grey
            
            Else
            
                Cells(i, 10).Interior.ColorIndex = 3 'negative/neutral=red
        
            End If
        
        Next i
Next

End Sub

