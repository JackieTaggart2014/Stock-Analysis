# VBA Challenge_ Module #2

## Overview of Project
### The purpose of this project was to provide analytical information to Steve and his parents regarding 12 green energy stocks. We looked at the data for 2017 and 2018 and provided information regarding each year’s stock performance. We then refactored the code in order to have it run more efficiently than the original code.

## Data 

### We were provided two charts with stock information on 12 different stocks. We looked at a few metrics to quantify our analysis such as entry date, Opening Value, Daily High Value, Daily Low Value, Closing Value, Adjusted Closing Value, and Volume traded. Each stock had a designated stock ticker, and each worksheet tracked the same stocks for each of the two years. We wanted to review the total daily volume traded for each stock ticker, and the yearly return on value. Once re ran the macros, we were given a run time to tell us how much time the macros took to run

![2017_OG Macros](https://user-images.githubusercontent.com/97082773/150021119-9e79c49e-1041-4b4b-83d8-f0b65a8e9287.png)
![2018_OG Macros](https://user-images.githubusercontent.com/97082773/150021149-d42c0b55-f9a3-4c1d-80d1-ee9cb7a1335e.png)


## Results

## Analysis

### In the original VBA macro, we were able to run code for each of the appropriate worksheets and created macros that would tell us how each stock was performing for years 2017 and 2018. We were then given additional criteria that could be entered into the code that would result in a more efficient code run.
       
       ### Additional Criteria:

  '1a) Create a ticker Index
   For i = 0 To 11
       tickerIndex = tickers(i)
       
       
    '1b) Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single, tickerEndingPrices As Single
       
       
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
       Worksheets(yearValue).Activate
       tickerVolumes = 0
       
       '2b) Loop over all the rows in the spreadsheet.
       
       For j = 2 To RowCount
              
           ' If the next row's ticker doesn't match, increase the tickerIndex.
           If Cells(j, 1).Value = tickerIndex Then
           
              '3a) Increase volume for current ticker
              tickerVolumes = tickerVolumes + Cells(j, 8).Value
        
           End If
           
           
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
           
           If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerStartingPrices = Cells(j, 6).Value
               
          'End If
           End If

        '3c) check if the current row is the last row with the selected ticker
        'If  Then
           
           If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerEndingPrices = Cells(j, 6).Value
               
          'End If
           End If
           
       Next j
       
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.

           Worksheets("All Stocks Analysis").Activate
           
           Cells(4 + i, 1).Value = tickerIndex
           Cells(4 + i, 2).Value = tickerVolumes
           Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
    
            'With Range("C4:C15")
                        '.NumberFormat = "0.0%"
                        '.Value = .Value
            'End With
            

   Next i

![VBA_ Challenge_2017_Run TIme](https://user-images.githubusercontent.com/97082773/150021924-40ffdd79-b14c-49f0-8e83-f53b06cd636a.png)
![VBA_ Challenge_2018_Run TIme](https://user-images.githubusercontent.com/97082773/150021940-6bc08991-073c-4f90-9371-482ac75f2298.png)

![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/97082773/150021979-2f747630-90a7-48ba-ac1b-59ca31d03cdc.png)
![VBA_Challenge_2018 png](https://user-images.githubusercontent.com/97082773/150021993-c0839138-e012-4703-b87b-3b4f48cd7350.png)


## Summary
	
## Advantages and disadvantages of refactoring code

Refactoring can make our code cleaner and easier to read. This could help in software improvement since the engineers will be able to decipher the code with greater ease. Refactoring can also help make the macro more efficient with shorter run times.
        The disadvantages are that sometimes with refactoring, the engineer can create a “fix” for the issue for that specific macro that cannot be transferred to another set of data. This could create issues since many engineers use their peers work and examples to base future code from. 
