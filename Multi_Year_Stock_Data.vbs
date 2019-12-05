Sub VBAChallenge():
    '   *********************************************************
    '   NOTE:  In the Script I assume that Data-Set was sorted by first column
    '   *********************************************************
    
    '   Variables:
    Dim ws As Worksheet
        
    Dim lRow As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim SumVolume As Double
    Dim TableRow As Long
    Dim TickerGInc As String
    Dim TickerGDec As String
    Dim TickerGVol As String
    Dim ValueGInc As Double
    Dim ValueGDec As Double
    Dim ValueGVol As Double
        
    ' Loop through all sheets

For Each ws In Sheets
    ws.Select
    
    
    '    Create the structure ot Table or Report 1:
    Cells(1, 9).Value = "Ticker"
    Cells(1, 9).Font.Bold = True
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 10).Font.Bold = True
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 11).Font.Bold = True
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 12).Font.Bold = True
   
    '    Create the structure ot Table or Report 2:
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(2, 15).Font.Bold = True
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(3, 15).Font.Bold = True
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(4, 15).Font.Bold = True
    Cells(1, 16).Value = "Ticker"
    Cells(1, 16).Font.Bold = True
    Cells(1, 17).Value = "Value"
    Cells(1, 17).Font.Bold = True
    
    '   *********************************************************
    '   Step 1:
    '   *********************************************************
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
    TableRow = 2
    OpenPrice = Cells(2, 3).Value
    ClosePrice = 0
    SumVolume = 0
   
    '    Loop through all Tickers with the same Label
   For i = 2 To lRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            Ticker = Cells(i, 1).Value
            
            ClosePrice = Cells(i, 6).Value
            
            SumVolume = SumVolume + Cells(i, 7).Value
            
            Range("I" & TableRow).Value = Ticker
            
            Range("J" & TableRow).Value = ClosePrice - OpenPrice
            
            If OpenPrice <> 0 Then
                Range("K" & TableRow).Value = (ClosePrice - OpenPrice) / OpenPrice
            Else
                Range("K" & TableRow).Value = 0
            End If
            
            Range("L" & TableRow).Value = SumVolume
            
            TableRow = TableRow + 1
            
            ' Initilalizing OpenPrice and SumVolume
             
            OpenPrice = Cells(i + 1, 3).Value
            
            SumVolume = 0
            
        Else
            
            SumVolume = SumVolume + Cells(i, 7).Value
                                    
        End If
                   
   Next i
    
    '   *********************************************************
    '   Step 2:
    '   *********************************************************
     
    For j = 2 To TableRow
    
        If Cells(j, 10).Value < 0 Then
        
        Cells(j, 10).Interior.ColorIndex = 3
        
        Else
        
        Cells(j, 10).Interior.ColorIndex = 4
        
        End If
    
    Next j

    '   *********************************************************
    '   Step 3:
    '   *********************************************************
        
        ValueGInc = Range("K2").Value
        ValueGDec = Range("K2").Value
        ValueGVol = Range("L2").Value
    
    For k = 2 To TableRow
       '***
        If Cells(k, 11).Value > ValueGInc Then
            
            TickerGInc = Cells(k, 9).Value
            ValueGInc = Cells(k, 11).Value
            
        End If
        '***
        If Cells(k, 11).Value < ValueGDec Then
            
            TickerGDec = Cells(k, 9).Value
            ValueGDec = Cells(k, 11).Value
            
        End If
        '***
        If Cells(k, 12).Value > ValueGVol Then
            
            TickerGVol = Cells(k, 9).Value
            ValueGVol = Cells(k, 12).Value
            
        End If
        
    Next k
    
    Range("P2").Value = TickerGInc
    Range("Q2").Value = ValueGInc
    Range("P3").Value = TickerGDec
    Range("Q3").Value = ValueGDec
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("P4").Value = TickerGVol
    Range("Q4").Value = ValueGVol
    
    '    *********************************************************
    '   Step 4:
    '   *********************************************************
     '  Adjust Columns
    
    Range("J2:J" & TableRow).NumberFormat = "0.00"
    
    Range("K2:K" & TableRow).NumberFormat = "0.00%"
    
    Columns("I:Q").AutoFit
    
    
Next ws

End Sub





