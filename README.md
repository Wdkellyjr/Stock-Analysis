# VBA Challenge – Green Stocks Analysis

## Overview of Project

A close friend of mine, Steve, asked if I could create an efficient way for him to analyze stock data from 2017 and 2018 in excel so that he could assist his parents in making the best decision in purchasing. Steve asked that the results be easy to read and accessible.

### Purpose

The purpose of the project is to use VBA script to efficiently analyze data from 2017 and 2018 for 12 stocks, so that Steve could see if the returns increased or decreased and the daily trading volume for each stock by year. Colors were used to highlight an increase in returns (green) and a decrease in returns (red). A “Run Analysis” button was added to make running the script easy for the user and a “Clear Cells” button was added so that the user could clear returned results quickly.

## Results

### Refactored Code Analysis

Since one of the main objectives of the project was to make a clean, efficient, VBA code, I refactored my original script to improve the design, structure, and implementation. I followed the instructions provided, used parts of my original code modified the for loops. 

### Refactored Code 

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
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next row’s ticker doesn’t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
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
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
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

## Summary

### Advantages and Disadvantages of Refactoring Code

The advantage of refactoring code is that it keeps the code clean, improves the performance of an application, saves time, money and reduces future technical expenses. The disadvantage of refactoring code is that it could be very time consuming depending on the project size and further mistakes can occur that will add additional time to the whole process. 

Refactoring the original script changed the runtime from about .45 seconds to about .12 seconds. The performance of the application was improved, and it took less time to run the code, which saves time and money. Since the code wasn’t extremely detailed, it didn’t take much time to make the improved changes.

### Refactored Code Runtimes

<img width="261" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/89553690/132960631-57a4de70-a50e-4464-860f-96aab91c16e4.png">

<img width="260" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/89553690/132960634-adf98574-eda0-494c-b899-0489f28631f3.png">
