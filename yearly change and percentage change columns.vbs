Attribute VB_Name = "Module3"
Sub yearly_change_and_Percentage()

'initiating script across worksheets

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    
ws.Activate
     
    'Create headers
    
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    
    'set up variables: op=first month/open price; cp=last month/close price; dl=tracker variable/earliest month; dh=tracker variable/last month
    
    Dim i, j As Long
    
    Dim op, cp, dl, dh As Long
    
    'set tracker variables to exceed values in date set
    
    dl = 33333333
    
    dh = 0
    
    j = 2
    
    'loops!
    
            'loop 1: evaluating closing price
            
            For i = 2 To Range("A2", Range("A1").End(xlDown)).Rows.Count + 1
                
                If Cells(i, 1).Value = Cells(j, 9).Value Then
                
                'evaluating closing price
                
                    If Cells(i, 2).Value > dh Then
            
                        cp = Cells(i, 6).Value
                        dh = Cells(i, 2).Value
                
                    Else
            
                    End If
                
                'evaluating opening price
        
                    If Cells(i, 2).Value < dl Then
            
                        dl = Cells(i, 2).Value
                        op = Cells(i, 3).Value
                
                    Else
            
                    End If
                
                Else
                
                'insert yearly change finding
                
                Cells(j, 10).Value = cp - op
                
                    If op = o Then
                    
                        Cells(j, 11).Value = "NaN"
                        
                    Else
                    
                    'insert yearly change finding
                    
                        Cells(j, 11).Value = FormatPercent((cp - op) / op)
                        
                    
                    End If
                       
                'resetting tracker calculations
                
                dh = o
            
                dl = Cells(i, 2).Value
                
                cp = 0
                op = Cells(i, 3).Value
            
                'move j counter
                
                j = j + 1
            
                End If
        
            Next i

Next

End Sub
