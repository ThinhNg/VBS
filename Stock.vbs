Sub Stock()
    'Written by Thinh Nguyen
    'November 11/18/2018
    'Completed 7:55PM
    Dim CurrentWb As Worksheet
    Dim TotalVol As Double
    Dim LastRow As Double
    Dim Summary_Row As Double
    Dim i As Double
    
    
    
    
    
    Dim OpenStock As Double
    Dim CloseStock As Double
    Dim YChange As Double
   
    
    
    
    Dim Max_PercentChange As Double
    Dim Low_PercentChange As Double
    Dim Max_TotalVol As Double
    
    Dim RowMaxP As Double
    Dim RowMinP As Double
    Dim RowMaxVol As Double
    
    'Total Vol is set to start at 0.
    TotalVol = 0
    'Iterates through each Worksheet.
    For Each CurrentWb In Worksheets
        
        'Sets headers for all necessary values for summary table.
        CurrentWb.Cells(1, 9).Value = "Ticker"
        CurrentWb.Cells(1, 10).Value = "Yearly Change"
        CurrentWb.Cells(1, 11).Value = "% Change"
        CurrentWb.Cells(1, 12).Value = "Total Stock Volume"
        
        CurrentWb.Cells(1, 15).Value = "Ticker"
        CurrentWb.Cells(1, 16).Value = "Value"
        CurrentWb.Cells(2, 14).Value = "Greatest % Increase"
        CurrentWb.Cells(3, 14).Value = "Greatest % Decrease"
        CurrentWb.Cells(4, 14).Value = "Greatest Total Volume"
        
        
        'Sets the last row of the current worksheet
        LastRow = CurrentWb.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Sets the starting point for the summary table for the tickers.
        Summary_Row = 2
        
        'Sets the Opening Stock value when transiting to a new worksheet.
        OpenStock = CurrentWb.Cells(2, 3).Value
       
        'Iterates until last row of current work book.
        For i = 2 To LastRow
            
            'Compares Current row ticker value with next contigent row ticker value to see if they are different.
            If CurrentWb.Cells(i, 1).Value <> CurrentWb.Cells(i + 1, 1).Value Then
                
                'Adds to the total Stock Volume for current ticker.
                TotalVol = TotalVol + CurrentWb.Cells(i, 7).Value
                
                'Sets Closing Stock for last day of current stock ticker.
                CloseStock = CurrentWb.Cells(i, 6).Value
                
                YChange = CloseStock - OpenStock
                
                'Sets Ticker Value in Column 9 and Total Stock Volume in column 10.
                CurrentWb.Cells(Summary_Row, 9).Value = CurrentWb.Cells(i, 1).Value
                CurrentWb.Cells(Summary_Row, 10).Value = YChange
                
                
                'Checks if Yearly Change is negative or green and colors it according.
                If YChange <= 0 Then
                    CurrentWb.Cells(Summary_Row, 10).Interior.ColorIndex = 3
                Else
                    CurrentWb.Cells(Summary_Row, 10).Interior.ColorIndex = 4
                End If
                    
                
                'Checks if there is any stock with No Change and 0 for opening stock value.
                If OpenStock = 0 Then
                   CurrentWb.Cells(Summary_Row, 11).Interior.ColorIndex = 5
                   CurrentWb.Cells(Summary_Row, 11).Value = 0
                   
                   
                Else
                    CurrentWb.Cells(Summary_Row, 11) = FormatPercent(YChange / OpenStock)
                End If
                CurrentWb.Cells(Summary_Row, 12) = TotalVol
                
                ' Resets Total Stock Volume to zero for a new stock Ticker.
                TotalVol = 0
                
                
                
                
                'Adds 1 to the Summary_Row for the next new Ticker.
                Summary_Row = Summary_Row + 1
                
                'Gets Starting Opening value of next stock.
                OpenStock = CurrentWb.Cells(i + 1, 3).Value
                
                
            'The Else indicates that the Ticker value for the next ticker is the same.
            Else
                'Adds to the total Stock Volume for current ticker.
                TotalVol = TotalVol + CurrentWb.Cells(i, 7).Value
            
            End If
        Next i
    
    '---------------------------------------------------------------------------
    'Add Summary of Greatest Change and Greatest % Total Increase and Decrease
    
    '----------------------------------------------------------------------------------
    
    'Sets for LastRow of Summary Table
    LastRow = CurrentWb.Cells(Rows.Count, 9).End(xlUp).Row
    
    
    'Below searches for the Max_PercentChange in the Summary Table and finds the row that it associated with.
    Max_PercentChange = WorksheetFunction.Max(CurrentWb.Range("K2:K" & LastRow)) * 100
    RowMaxP = CurrentWb.Range("K2:K" & LastRow).Find(Max_PercentChange).Row
    
    'Below searches for the Min_PercentChange in the Summary Table and finds the row that it associated with.
    Min_PercentChange = WorksheetFunction.Min(CurrentWb.Range("K2:K" & LastRow)) * 100
    RowMinP = CurrentWb.Range("K2:K" & LastRow).Find(Min_PercentChange).Row
    
    
    'Below searches for the maximum total stock volume in the Summary Table and finds the row that it associated with.
    Max_TotalVol = WorksheetFunction.Max(CurrentWb.Range("L2:L" & LastRow))
    RowMaxVol = CurrentWb.Range("L2:L" & LastRow).Find(Max_TotalVol).Row
    

    'The following sets up the summary table values for the greatest % increase/decrease and the maxmimum total volume.
    CurrentWb.Cells(2, 15) = CurrentWb.Cells(RowMaxP, 9)
    CurrentWb.Cells(2, 16) = FormatPercent(CurrentWb.Cells(RowMaxP, 11))
    
    CurrentWb.Cells(3, 15) = CurrentWb.Cells(RowMinP, 9)
    CurrentWb.Cells(3, 16) = FormatPercent(CurrentWb.Cells(RowMinP, 11))
    
    CurrentWb.Cells(4, 15).Value = CurrentWb.Cells(RowMaxVol, 9).Value
    CurrentWb.Cells(4, 16).Value = CurrentWb.Cells(RowMaxVol, 12).Value
    
    
    
    
    'Iterates to next workbook
    Next CurrentWb
        
End Sub
            
            
            
                
                
                
                
                
                
                
    
    
    
    





