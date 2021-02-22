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

### Run-Time Comparison
Original Code: 2017

Original Code: 2018

Refactored Code: 2017

Refactored Code: 2018

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
