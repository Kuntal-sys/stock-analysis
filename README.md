# stock-analysis
Sub Challenge2()
    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("Challenge2_All Stocks Anlysis").Activate
    
    Range("A1").Value = "Stocks analysis for (" + yearValue + ")"
    
    'STEP 1:Creating header rows
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'STEP 2:declaring all arrays
    Dim tickers(12) As String
    Dim volume(12) As String
    Dim startPrice(12) As String
    Dim endPrice(12) As String
    
    
    'STEP 3:creating index variable
    Dim tickerIndex As Integer

    Worksheets(yearValue).Activate
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'STEP 4:creating outer loop for index from 0 to 11
    tickerIndex = 0
    Worksheets(yearValue).Activate
    
    For tickerIndex = 0 To 11
        
        
    'STEP 5:creating loop for storing all stock data for selected year
        
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
            
     'STEP 6:Fetch ticker name and start price for each and store them in arrays
            
            If Cells(j, 1).Value <> Cells(j - 1, 1).Value Then
                tickers(tickerIndex) = Cells(j, 1).Value
                startPrice(tickerIndex) = Cells(j, 6).Value
            
            End If
                
      'STEP 7:creating nested loop for feteching TotalVolume for each volume of array
                
                Worksheets(yearValue).Activate
                    TotalVolume = 0
                    For x = 2 To RowCount
                        If Cells(x, 1).Value = tickers(tickerIndex) Then
                            TotalVolume = TotalVolume + Cells(x, 8).Value
                        End If
                    Next x

                    volume(tickerIndex) = TotalVolume
            
            'STEP 8:fetching end price and storing in array
            
            If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
                endPrice(tickerIndex) = Cells(j, 6).Value
                
                'Trigger tickerIndex for next loop
                tickerIndex = tickerIndex + 1
                
            End If
            
        Next j
        
    Next tickerIndex
    
    'STEP 9: storing all fetched information from array and place it in below worksheet
    Worksheets("Challenge2_All Stocks Anlysis").Activate
    For i = 0 To 11
        
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 3).Value = endPrice(i) / startPrice(i) - 1
        Cells(4 + i, 2).Value = volume(i)
    
    Next i
                
    'STEP 10: formatting
    Worksheets("Challenge2_All Stocks Anlysis").Activate
        Range("A3:C3").Font.Bold = True
        Range("A1").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns(2).AutoFit
    
    
    'STEP 11: conditional/color formatting
    Worksheets("Challenge2_All Stocks Anlysis").Activate
    dataRowEnd = Cells(Rows.Count, "C").End(xlUp).Row
    dataRowStart = 4
    
    For y = dataRowStart To dataRowEnd
        
        If Cells(y, 3).Value > 0 Then
           Cells(y, 3).Interior.Color = vbGreen
        
        ElseIf Cells(y, 3).Value < 0 Then
               Cells(y, 3).Interior.Color = vbRed
        
        Else
            Cells(y, 3).Interior.Color = xlNone
       
        End If
    
    Next y


End Sub

