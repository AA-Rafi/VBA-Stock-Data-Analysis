Sub challenge1():
    
    
    Dim lastrow As Variant
    Dim name As String
    Dim sumrow As Variant
    Dim vol As Variant
    Dim firstval As Variant
    Dim lastval As Variant
    Dim yearchange As Variant
    
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    sumrow = 2

    'Setting up column names
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Year Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

    
    For i = 2 To lastrow

    'Ticker and Volume column based heavily off of the CreditCard() exercise

        If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
        
            name = Cells(i, 1).Value
            vol = vol + Cells(i, 7).Value
        
            Range("I" & sumrow).Value = name
            Range("L" & sumrow).Value = vol
        
            firstval = Cells(Application.Match(name, Range("A:A"), 0), 3)
            lastval = Cells(Application.Match(name, Range("A:A"), 0), 6)
            'LASTVAL is clearly WRONG, using the first value in F rather than the last and I cant figure out how to fix it
        
            yearchange = lastval - firstval
            Range("J" & sumrow).Value = yearchange
        
            perchange = (yearchange / firstval) * 100
            Range("K" & sumrow).Value = perchange
            'Changing Percent change into percent!
            Range("K" & sumrow).NumberFormat = "0.00%"
            sumrow = sumrow + 1
            vol = 0
        Else
        
        vol = vol + Cells(i, 7).Value
      
    End If

    'Year change color input
    
        If (Cells(i, 10) > 0) Then
            Cells(i, 10).Interior.ColorIndex = 4
    
        Else
            Cells(i, 10).Interior.ColorIndex = 3
        End If

Next i

End Sub

Sub challenge2():
    
    Dim MaxChange As Variant
    
    Dim MinChange As Variant
    
    Dim MaxVol As Variant
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    sumrow = 2
    
    For i = 2 To lastrow
    
    'Table Names Setup
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
        
        
    'Greatest % Increase
        If Cells(i, 11).Value = WorksheetFunction.Max(Range("K2:K" & sumrow)) Then
                
            MaxChange = WorksheetFunction.Max(Range("K2:K" & lastrow))
            Range("Q2").Value = MaxChange
            Range("Q2").NumberFormat = "0.00%"
            
            Range("P2").Value = Cells(Application.Match(MaxChange, Range("K:K"), 0), 9)
                
        End If
    
    'Greatest % Decrease
        If Cells(i, 11).Value = WorksheetFunction.Min(Range("K2:K" & sumrow)) Then
        
            MinChange = WorksheetFunction.Min(Range("K2:K" & lastrow))
            Range("Q3").Value = MinChange
            Range("Q3").NumberFormat = "0.00%"
        
            Range("P3").Value = Cells(Application.Match(MinChange, Range("K:K"), 0), 9)
          
        End If

    'Greatest Total Volume
        If Cells(i, 12).Value = WorksheetFunction.Max(Range("L2:L" & sumrow)) Then
                
            MaxVol = WorksheetFunction.Max(Range("L2:L" & lastrow))
            Range("Q4").Value = MaxVol
    
            Range("P4").Value = Cells(Application.Match(MaxVol, Range("L:L"), 0), 9)
        
        End If
    
    Next i

End Sub

