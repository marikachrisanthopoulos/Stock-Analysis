# Stock-Analysis
## Project Overview
The purpose of this project was to analyze stock data for my friend Steve, a newly graduated finance major. Steve's parents are passionate about green energy, and have decided to invest all of their money into DAQO New Energy Corp, a company that makes silicon wafers for solar panels. I used Visual Basic for Applications (VBA) in Microsoft Excel to analyze this stock's total daily volume and annual return. Next, I analyzed 11 additional green stocks to determine whether DAQO is their best option to invest in, based on these statistics. The purpose of using VBA was to create an efficient method for analyzing the data; the focus of this challenge was to restructure the code so it could be easily replicated for future analyses in a process known as refactoring.
## Results
Refactoring the code consisted of completing the following steps:
- Creating a ticker index variable (tickerIndex) to access the correct index across four different arrays.
- Creating three arrays (tickerVolumes, tickerStartingPrices, tickerEndingPrices)
- Creating a loop through the stock data to read and store the following values from each row: tickers, tickerVolumes, tickerStartingPrices, tickerEndingPrices.

See the refactored code below:

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
        tickerVolumes(tickerIndex) = 0
        
    ''2b) Loop over all the rows in the spreadsheet.
        For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then

        If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row‚Äôs ticker doesn‚Äôt match, increase the tickerIndex.
        'If  Then
            
             If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

            '3d Increase the tickerIndex.
            
            tickerIndex = tickerIndex + 1
        
        End If
        
        Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = TotalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

    Next i
     
    End Sub


## Summary
There are advantages and disadvantages to refactoring code, as evidenced by this project - see below:
### Advantages
- Potential for greater efficiency (faster; less code for the same amount of tasks; easily adjustable for shifting variables, such as year) 
- Potential to be more easily read/understood by other programmers
- Potential to be used for similar future analyses (easier template to edit and reuse)
- Original code can be used and then edited (programmers do not need to start from scratch)
### Disadvantages
- Potential of disrupting the code to make it unuseable (until debugging procedures figure out the error)
### Overall
In general, I found refactoring the code to be a difficult, but rewarding, task. Considering my understanding of the syntax is still limited, I had a challenging time working through the errors I created while trying to make the code more efficient. However, working through these errors was helpful to enhance my knowledge of VBA syntax, and felt like an accomplishment once I finally was able to successfuly run the task.
