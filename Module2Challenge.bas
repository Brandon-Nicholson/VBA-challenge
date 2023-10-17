Attribute VB_Name = "Module6"
Sub SheetLoop()

Dim ws As Worksheet

' loop through worksheets
For Each ws In Worksheets
    
    'Create column names with bold format
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
    
    Range("I1:L1").Select
    
    Selection.Font.Bold = True
    
    Range("A1:Q100000").Columns.AutoFit
    
    
    'Fill column with all uniquw tickers
    
    counter = 1
    
    For r = 2 To Range("A" & Rows.Count).End(xlUp).Row
    
        If Cells(r, 1) <> Cells(r + 1, 1) Then
            counter = counter + 1
            Cells(counter, 9) = Cells(r, 1)
            
        End If
        
    Next r
    
    
    'Subtract yearly open - close and fill Yearly Change column
    
    counter2 = 1
    open1 = 0
    close1 = 0
    
    For a = 2 To Range("A" & Rows.Count).End(xlUp).Row
        
        
        If Cells(a, 1) <> Cells(a - 1, 1) Then
            counter2 = counter2 + 1
            
            open1 = Cells(a, 3)
            
        End If
            
        If Cells(a, 1) <> Cells(a + 1, 1) Then
            close1 = Cells(a, 6)
                
            Cells(counter2, 10) = (open1 - close1) * -1
                
        End If
        
    Next a
        
        
    'Percentage change from opening price to closing price
        
    counter3 = 1
    open2 = 0
    close2 = 0
    
    For s = 2 To Range("A" & Rows.Count).End(xlUp).Row
        
        
        If Cells(s, 1) <> Cells(s - 1, 1) Then
            counter3 = counter3 + 1
            
            open2 = Cells(s, 3)
            
        End If
            
        If Cells(s, 1) <> Cells(s + 1, 1) Then
            close2 = Cells(s, 6)
            
            Change = Round(((open2 - close2) * -1) / open2 * 100, 2)
            
            Cells(counter3, 11) = Str(Change) + "%"
            
            If Change < 0 Then
            
                Cells(counter3, 11).Interior.ColorIndex = 3
            
            Else
            
                Cells(counter3, 11).Interior.ColorIndex = 4
            
            End If
                
        End If
        
    Next s
    
    
    'Total volume of each stock
    
    counter4 = 1
    volume = 0
    
    For q = 2 To 753002
        
        volume = volume + Cells(q, 7)
        
        If Cells(q, 1) <> Cells(q + 1, 1) Then
            counter4 = counter4 + 1
            Cells(counter4, 12) = volume
            volume = 0
            
        End If
                          
        
    Next q
    
    
    'Show greatest increase in volume
    
    increase = -100
    ticker = 0
    
    For e = 2 To Range("I" & Rows.Count).End(xlUp).Row
    
        If Cells(e, 11) > increase Then
            increase = Cells(e, 11)
            ticker = Cells(e, 9)
        End If
    
    Next e
    
    
    Range("P2") = ticker
    Range("Q2") = Str(increase * 100) + "%"
    
    
    'Show greatest decrease in volume
    
    decrease = 100
    ticker2 = 0
    
    For w = 2 To Range("I" & Rows.Count).End(xlUp).Row
    
        If Cells(w, 11) < decrease Then
            decrease = Cells(w, 11)
            ticker2 = Cells(w, 9)
        End If
    
    Next w
    
    
    Range("P3") = ticker2
    Range("Q3") = Str(decrease * 100) + "%"
    
    
    'Greatest total volume
    
    volume = 0
    ticker3 = 0
    
    For v = 2 To Range("I" & Rows.Count).End(xlUp).Row
    
        If Cells(v, 12) > volume Then
            volume = Cells(v, 12)
            ticker3 = Cells(v, 9)
        End If
    
    Next v
    
    
    Range("P4") = ticker3
    Range("Q4") = volume
    
    ws.Activate


Next

End Sub
