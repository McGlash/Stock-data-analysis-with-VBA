Sub Ticker_column_RUN_FIRST()

'! must be run first

'initiating script across worksheets

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
    
 ws.Activate
 
 'sort
 
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.sort key1:=Range("A1", Range("A1").End(xlDown)), order1:=xlAscending, key2:=Range("B1", Range("B1").End(xlDown)), order2:=xlAscending, Header:=xlYes

    ' set up header
    
    Cells(1, 9).Value = "Ticker"

    Dim i, j As Long
    
    'set up tracker variable

    j = 2

    'set start of zero value
    
    Cells(2, 9).Value = Cells(2, 1).Value

    'loop!
    
    For i = 2 To Range("A2", Range("A1").End(xlDown)).Rows.Count

        If Cells(i, 1).Value <> Cells(j, 9).Value Then
            j = j + 1
            Cells(j, 9).Value = Cells(i, 1).Value
        
    Else
    
    End If
    
    Next i
    
Next

End Sub

