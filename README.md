# Stock analysis with Excel
A client wanted to analyze the stock market over the last few years.  This project required refactoring VBA code to loop through data once in order to collect the same information.    


## Overview of Project
Using a stock market dataset spanning two years, VBA code was written to calculate total daily volume and rate of return for a year.  The stock market data was contained in an Excel spreadsheet, and a VBA script was written.  The VBA code was refactored in order to make it more efficient to use against larger datasets.


## Analysis
- The original VBA code included two loops:
 
 '4. Loop through all tickers (AKA 'ticker loop')

For i = 0 To 11
        
    ticker = tickers(i)
    totalVolume = 0


    '5. Loop through rows in the data (AKA 'row loop')

    Worksheets(yearValue).Activate
     For j = 2 To RowCount


        '5a. Find the total volume for the data
    
        If Cells(j, 1).Value = ticker Then
        
         totalVolume = totalVolume + Cells(j, 8).Value
        
         End If
    
    
         '5b. Find the starting price for the current ticker using multiple conditions
    
          If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
             startingPrice = Cells(j, 6).Value
        
         End If
    
    
        '5c. Find the ending price for the current ticker
    
         If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
    
            endingPrice = Cells(j, 6).Value
        
            End If
        
        Next j

   * When this code was executed against the 2017 stock market dataset, it ran in 0.578 seconds.
   
   ![](Resources/Module_2_2017.png)
   
   
   * When this code was executed against the 2018 stock market dataset, it ran in 0.585 seconds.
   
   ![](Resources/Module_2_2018.png)
   
   
- The refactored VBA code was consolidated into one loop:

    For i = 2 To RowCount
    
    
        '3a) Increase volume for current ticker.
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.  This finds the ticker's starting price.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
            
        
        '3c) check if the current row is the last row with the selected ticker. If the next row’s ticker doesn’t match, increase the tickerIndex. This finds the ticker's ending price.
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
      
            
            '3d Increase the tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
                
            End If
            
    Next i
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        'Output for each ticker will print in a new row.  Value will print on 4th row + i.
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
	
	
   * When the refactored code was executed against the 2017 stock market data set, it ran in 0.128 seconds, which is 4.5 times faster than the original code.
   
   ![](Resources/VBA_Challenge_2017.png)
   
   
   * When the refactored code was executed against the 2018 stock market dataset, it ran in 0.101 seconds.  This is nearly 6 times faster than the original code.
   
   ![](Resources/VBA_Challenge_2018.png)


## Summary
- What are the advantages of refactoring code?
   * Increases speed of program
   * Makes the code easier to understand
   * Helps find bugs
   
   
- What are the disadvantages of refactoring code?
   * The developer must understand what the code is doing in order to improve upon it
   * Refactoring code could produce bugs
   * It could be risky in terms of time and money - "if it ain't broke, don't fix it"


- How do these pros and cons apply to refactoring the original VBA script?
   * The refactored VBA code increased in execution speed
   * The initial code was easy to understand because I wrote it
   * The initial VBA code seemed stable, so in terms of spending time and money to refactor it, it may not be worthwhile
   
   