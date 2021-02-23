# Stock-Analysis
## Project Overview
The purpose of this project was to analyze stock data for my friend Steve, a newly graduated finance major. Steve's parents are passionate about green energy, and have decided to invest all of their money into DAQO New Energy Corp, a company that makes silicon wafers for solar panels. I used Visual Basic for Applications (VBA) in Microsoft Excel to analyze this stock's total daily volume and annual return. Next, I analyzed 11 additional green stocks to determine whether DAQO is their best option to invest in, based on these statistics. The purpose of using VBA was to create an efficient method for analyzing the data; the focus of this challenge was to restructure the code so it could be easily replicated for future analyses in a process known as refactoring.
## Results
Refactoring the code consisted of completing the following steps:
- Creating a ticker index variable (tickerIndex) to access the correct index across four different arrays.
- Creating three arrays (tickerVolumes, tickerStartingPrices, tickerEndingPrices)
- Creating a loop through the stock data to read and store the following values from each row: tickers, tickerVolumes, tickerStartingPrices, tickerEndingPrices.

Implementing the tickerIndex variable allowed me to analyze the data faster than before: this is because the other three arrays (tickerVolumes, tickerStartingPrices, and tickerEndingPrices) could be assigned to each ticker (stock) before iterating through the data set.

See the refactored code below:

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
        Dim tickerIndex As Single
        tickerIndex = 0
        
    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        tickerVolumes(i) = 0
        Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then

        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
        '3c) check if the current row is the last row with the selected ticker
            'If the next row's ticker doesn't match, increase the tickerIndex.
            'If  Then
            
             If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
        
        End If
        
        Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(i + 4, 1).Value = tickers(tickerIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
     
     Next i
     
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
     
    End Sub

In addition, the results both both 2017 and 2018 stocks are below:

![2017 Results](https://github.com/marikachrisanthopoulos/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017_Results.png)

![2018 Results](https://github.com/marikachrisanthopoulos/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018_Results.png)

### Run-Time Comparison

The time it takes to run the analysis via the original code (2017, 2018) and the refactored code (2017, 2018) is shown below. There is approximately a second difference between the original and the refactored code for both years.

Original Code: 2017
![2017 Original Time](https://github.com/marikachrisanthopoulos/Stock-Analysis/blob/main/Resources/VBA_PreRefactored_2017.png)

Original Code: 2018
![2018 Original Time](https://github.com/marikachrisanthopoulos/Stock-Analysis/blob/main/Resources/VBA_PreRefactored_2018.png)

Refactored Code: 2017
![2017 Refactored Time](https://github.com/marikachrisanthopoulos/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)

Refactored Code: 2018
![2018 Refactored Time](https://github.com/marikachrisanthopoulos/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary
There are advantages and disadvantages to refactoring code, both in general and in this project specifically - see below:
### Advantages - General
- Potential for greater efficiency (faster; less code for the same amount of tasks; easily adjustable for shifting variables)
- Potential to be more easily read/understood by other programmers
- Potential to be used for similar future analyses (easier template to edit and reuse)
- Original code can be used and then edited (programmers do not need to start from scratch)
### Disadvantages - General
- Potential of disrupting the code to make it unuseable (until debugging procedures figure out the error)
### Advantages - Stock Analysis Project
- Greater efficiency (over a second faster than the original code; easily adjustable for shifting variables)
- Can be easily interpreted/edited by other programmers for future analyses
### Disadvantages - Stock Analysis Project
- The original code was disrupted and took time to reformulate
### Overall
In general, I found refactoring the code to be a difficult, but rewarding, task. Considering my understanding of the syntax is still limited, I had a challenging time working through the errors I created while trying to make the code more efficient. However, working through these errors was helpful to enhance my knowledge of VBA syntax, and felt like an accomplishment once I finally was able to successfuly run the task.
