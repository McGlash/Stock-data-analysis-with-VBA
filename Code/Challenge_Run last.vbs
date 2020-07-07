Attribute VB_Name = "Module6"
Sub summary()

'warning: this script must be run last.

'script across sheets

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    
ws.Activate

'layout table

Range("N2") = "Greatest % Increase"
Range("N3") = "Greatest % Decrease"
Range("N4") = "Greatest Total Volume"

Range("O1") = "Ticker"
Range("P1") = "Value"

'set up tracker variables

''volume
Dim vol As Double

vol = Cells(2, 12).Value

''Greatest increase
Range("P2").Value = Cells(2, 11).Value

''greatest decrease
Range("P3").Value = Cells(2, 11).Value

    'Loops!
    
    Dim i As Integer
    
    For i = 3 To Range("I2", Range("I1").End(xlDown)).Rows.Count + 1
    
    'greatest decrease loop

        If Cells(i, 11).Value <> "NaN" Then
        
            If Cells(i, 11).Value < Range("P3").Value Then
            
            Range("P3").Value = Cells(i, 11).Value
            Range("O3").Value = Cells(i, 9).Value
            
            Else
            
            End If
        
            If Cells(i, 11).Value > Range("P2").Value Then
            
            Range("P2").Value = Cells(i, 11).Value
            Range("O2").Value = Cells(i, 9).Value
            
            Else
            
            End If
             
        Else
            
        End If
       
       'greatest volume loop
    
        If Cells(i, 12).Value > vol Then
        
            vol = Cells(i, 12).Value
            
            vol_t = Cells(i, 9).Value
            
        Else
        
        End If
    
    Next i

    'Insert volume findings
    
    Range("O4") = vol_t
    
    Range("P4") = vol
    

Next

End Sub
