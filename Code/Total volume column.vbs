Attribute VB_Name = "Module2"

Sub volume()

'initiating script across worksheets

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    
ws.Activate

    'set up header
    
    Cells(1, 12).Value = "Total Volume"
    
    Cells(2, 12).Value = Cells(2, 7).Value
    
    'Set up loop variables
    
    Dim i, j As Long
    
    j = 2
    
    'loop!
    
        For i = 3 To Range("A2", Range("A1").End(xlDown)).Rows.Count + 1
            
            If Cells(i, 1).Value = Cells(j, 9).Value Then
            
                Cells(j, 12).Value = Cells(j, 12).Value + Cells(i, 7).Value
                
            Else
            
            Cells(j + 1, 12).Value = Cells(i, 7).Value
            
            j = j + 1
            
            
            End If
            
        Next i
    
Next

End Sub


