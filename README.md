# Stock Analysis
This report is an analysis into the stock market from the years **2017 to 2018**. We will be looking into the total volume traded for our stocks as wll as their returns for their respective years.

## Overview of Project
Our client Steve wants us to look into the data and help him make an informed decision as to which stocks to invest in.

### Purpose
We have to fullfill two goals for this report
* Recommend Steve what stocks to buy so he can advice his parents
* To refactor an exsisting VBA code for a similar purpose as such the new code will be able to work with a larger dataset. As well as being more efficient and less time consuming.

## Results
Lets take a closer look at the analysis and our findings.

### Analysis
Our dataset consists of two sheets (_Year 2017 and 2018_) with information regarding 12 stocks. 

First we organized the data and extracted three output coloumns on a seperate sheet (_All Stocks Analysis_).
* Ticker
* Total Daily Volume
* Return

We had two tables for our findings, one for the year 2017 and the second for the year 2018.

![Screenshot 2021-11-11 145444](https://user-images.githubusercontent.com/93144225/141360693-9b014bb9-fe03-4847-ac8d-a9f7b9f8b6bb.png)

![201111](https://user-images.githubusercontent.com/93144225/141360581-aca5cdb3-a3c2-46eb-bc61-794159fc1fbc.png)

In order to get our results, we refactored an existing VBA code. I have attached the code below with helpfull comments to clarify and explain what line is doing what.

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
    Dim tickers(11) As String
    
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
    RowStart = 2
    
    '1a) Create a ticker Index

    tickerindex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(11) As Long
    
    Dim tickerStartingPrices(11) As Single
    
    Dim tickerEndingPrices(11) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        
    Next i
   
    ''2b) Loop over all the rows in the spreadsheet.
    For i = RowStart To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
            
            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            
            tickerEndingPrices(tickerindex) = Cells(i, 6).Value
         
         End If

        '3d) Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
                
                tickerindex = tickerindex + 1
            
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

```

## Summary

### Why refactor code?
There is a lot of pros and cons when it comes to refactoring code.

#### Pros
* The code will be cleaner and more organized. As it is a revision of an existing code, a lot of things that were overlooked the first time would be fixed.
* The code will be more efficient as a lot of extra or unnecessary steps will be deleted. This will make the code more concise.
* The code will be more easy to understand as comments can be added given the second revision.

#### Cons
* If we are dealing with huge codes, it is a significant time and developer work investment. Sometimes the extra efficiency and cleaner code is not worth that extra investment.
* For very complex codes, if the refactoring is not done carfully a lot of bugs might show up. Causing more harm then good.

### Why refactor code for Stock Analysis?
* The refactored code runs more efficiently then our older code and is much faster when it comes to macro run time.
* The new code can handle a lot more data then the older one. The previous code only handled one stock whereas the new one handles much more.
* The newer code has visual indicators (_Positive returns are highlighted in green and negetive ones are red_) whereas the older one was plain. This makes the newer worksheet more easy to understand.
* The older code had a smaller file size then the newer one.

## Links
  * Visit this [link](https://github.com/tanzimamin2/stock-analysis) for the excel worksheet and other resources.
   
