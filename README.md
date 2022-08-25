# stocks-analysis

## Overview of Project
This project was meant to analyze stock information between 2017 and 2018 through the refactoring of a Microsoft Excel VBA code in order to further evaluate whether to invest in certain stocks or not. An initial code was created in order to provide the data analysis; however, this project also allowed for more efficiency and clearer data formatting.

## Results
For this project, I entered the provided code into the VBA editor:
```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    

    '1b) Create three output arrays   
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    
        
    ''2b) Loop over all the rows in the spreadsheet. 
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            

            '3d Increase the tickerIndex. 
            
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

In order to refactor the code, I began with step 1a, in which I created a tickerIndex and set it to zero. 
![tickerIndex](https://user-images.githubusercontent.com/110862583/186558075-91f7a66a-f82e-4c8f-9b47-0948878e46ba.png)

From there, I created three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices and set the arrays as Long and Single data types. 
![Arrays](https://user-images.githubusercontent.com/110862583/186558287-8114904e-a184-467e-95f2-a83f1c9ffc3d.png)

I then entered the following code to initialize the tickerVolumes at 0 and loop over the rows in the respective worksheet (which would be based on the desired year):
![Initial and loop](https://user-images.githubusercontent.com/110862583/186558500-77707b45-c5b0-4fa3-abc6-5784a8d191b9.png)

In order to increase the volume for the current ticker, I entered
```
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
```
using tickerIndex as the variable.

In order to check which rows were selected, as well as to increase the tickerIndex based off the results, I used the following codes:
![First, last, and increase](https://user-images.githubusercontent.com/110862583/186567715-b93bf375-46f2-4da7-b0aa-a5d7a36dc51a.png)

Finally, to display the output of the tickers, Total Daily Volume, and Return, I entered: 
![Output](https://user-images.githubusercontent.com/110862583/186567906-5cc9b1f7-15c0-4390-a607-b958898c5e7e.png)

## Summary
### Pros and Cons of Refactoring Code

Refactoring code can allow for a more efficient and more organized code when analyzing large datasets, which may prove beneficial for individuals viewing our projects and datasets. Through refactoring, we can also provide datasets with a formatted design and debug issues in our code. However, due to the size and risk of running large macros onexcel, refactored code may prove to be detrimental when sharing to other users. 

### Advantages and Disadvantages of Refactored Code Script

The most prominent advantage of refactored code is the run time. The original code took longer to run -- approximately 1-1.2 seconds. The refactored code processed the data at nearly 1/4 of the run time for both years: 
![2018](https://user-images.githubusercontent.com/110862583/186569276-d881742c-52ad-4214-99c5-70abf01b96e9.png)
![2017](https://user-images.githubusercontent.com/110862583/186569277-9219d7ce-4bed-48d7-b69d-99c872ee18f6.png)
