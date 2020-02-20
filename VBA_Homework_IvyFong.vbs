Sub VBA_Stocks()

    ' Instructions: Create a script that will loop through all the stocks for one year for each run and take the following information.
        ' The ticker symbol.
        ' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
        ' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        ' The total stock volume of the stock.
        ' You should also have conditional formatting that will highlight positive change in green and negative change in red.
    
    ' Challenge
        ' Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
        ' Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.
    
    
    
    ' Run on every worksheet - From 1 to last worksheet
    
    For Index = 1 To Sheets.Count
    
    
    ' Activate next worksheet
    
    Worksheets(Index).Activate
    
    
        ' ASSIGNMENT
        
        
        ' Add labels
        
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        
        ' Find last row of column A
        
        ColALastRow = Cells(Rows.Count, 1).End(xlUp).Row
            
            
        ' Fill in column I with ticker values
            ' Start with ColIRow = 2 and PrevTicker = "Z"
            ' For i = 2 to last row number in column A
                ' CurrentTicker = Value of cells in column A
                    ' If CurrentTicker does not equal to PrevTicker Then
                    ' Fill in column I with value of CurrentTicker
                    ' Add 1 to variable that specifies the next empty row in column I (ColIRow)
        
        ColIRow = 2
        PrevTicker = "Z"
        
        For i = 2 To ColALastRow
            
            CurrentTicker = Cells(i, 1).Value
            
            If CurrentTicker <> PrevTicker Then
                Cells(ColIRow, 9).Value = CurrentTicker
                PrevTicker = CurrentTicker
                ColIRow = ColIRow + 1
            End If
            
        Next i


        ' Fill in columns J-L with the appropriate values
            ' Start with OpenRow = 2 and ColIRow = 2
            ' For i = OpenRow to last row number in column A + 1
            ' CurrentTicker = value of cells in column A
            ' Ticker = Value of cells in column I
                
                ' If CurrentTicker does not equal to Ticker Then
                ' Set CloseRow = i - 1
                
                ' Calculate yearly change value and fill in column J with the value
                ' Calculate percent change value and fill in column K with the value
                
                ' Add 1 to variable that specifies the next empty row in column I (ColIRow)
                ' Specify that i = first row for the next ticker
        
        ' OpenRow = Row of first occurence of ticker = Opening date for ticker
        ' CloseRow = Row of last occurence of ticker = Closing date for ticker
        
        OpenRow = 2
        ColIRow = 2
        
        For i = OpenRow To (ColALastRow + 1)
            
            CurrentTicker = Cells(i, 1).Value
            Ticker = Cells(ColIRow, 9).Value
            
            If CurrentTicker <> Ticker Then
                    
                    CloseRow = i - 1
                    
                    YearlyChange = Cells(CloseRow, 6).Value - Cells(OpenRow, 3).Value
                    Cells(ColIRow, 10).Value = YearlyChange
                    
                    If Cells(OpenRow, 3).Value <> 0 Then
						PercentChange = YearlyChange / Cells(OpenRow, 3).Value
						Cells(ColIRow, 11).Value = PercentChange
						Cells(ColIRow, 11).NumberFormat = "##0.00%"
					Else 
						PercentChange = "N/A"
                    End If
                    
                    TotalStockVolume = WorksheetFunction.Sum(Range(Cells(OpenRow, 7), Cells(CloseRow, 7)))
                    Cells(ColIRow, 12).Value = TotalStockVolume
                    
                    ColIRow = ColIRow + 1
                    OpenRow = i
            
            End If
            
        Next i


        ' Find last row of column J
        
        ColJLastRow = Cells(Rows.Count, 10).End(xlUp).Row
        
        
        ' Define colours
        
        Green = RGB(0, 255, 0)
        Red = RGB(255, 0, 0)
        
        
        ' Conditioning formatting to highlight positive yearly change in green and negative in red
            ' For i = 2 to last row in column J
            ' If yearly change is greater than 0, highlight cell in green
            ' Else if yearly change is less than 0, highlight cell in red
            
        For i = 2 To ColJLastRow
            
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.Color = Green
            ElseIf Cells(i, 10).Value < 0 Then
                Cells(i, 10).Interior.Color = Red
            End If
            
        Next i
        
        
        
        ' CHALLENGE
        
        
        ' Add labels
        
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        
                
        ' Fill in columns O-P with the appropriate values
            ' Find the max value in column K and output to the greatest % increase value
            ' Find the min value in column K and output to the greatest % decrease value
            ' Find the max value in column L and output to the greatest total volume value
            ' For i = 2 to last row number in column J (same number of rows as columns K and L)
                ' Fill in column O with the corresponding ticker
            
        GreatestPercentIncrease = WorksheetFunction.Max(Range(Cells(2, 11), Cells(ColJLastRow, 11)))
        Range("P2") = GreatestPercentIncrease
        
        GreatestPercentDecrease = WorksheetFunction.Min(Range(Cells(2, 11), Cells(ColJLastRow, 11)))
        Range("P3") = GreatestPercentDecrease
        
        GreatestTotalVolume = WorksheetFunction.Max(Range(Cells(2, 12), Cells(ColJLastRow, 12)))
        Range("P4") = GreatestTotalVolume
        
        For i = 2 To ColJLastRow
            
            If Cells(i, 11).Value = GreatestPercentIncrease Then
                Range("O2") = Cells(i, 9).Value
				Cells(2,16).NumberFormat = "##0.00%"
            End If
            
            If Cells(i, 11).Value = GreatestPercentDecrease Then
                Range("O3") = Cells(i, 9).Value
            End If
            
            If Cells(i, 12).Value = GreatestTotalVolume Then
                Range("O4") = Cells(i, 9).Value
				Cells(3,16).NumberFormat = "##0.00%"
            End If
            
        Next i
        
        
        ' Autofit columns
        
        Columns("I:L").AutoFit
        Columns("N:P").AutoFit
    
    
    Next Index
    
    
End Sub
            