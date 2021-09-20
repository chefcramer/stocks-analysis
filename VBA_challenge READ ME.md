# VBA of Wallstreet

## Overview of Project

This project is to analyze the overall performance of a data set of 12 different "Green" companies to determine which is the best investment opportunity. The Data was organized, analyzed and formatted using VBA (Visual Basic), a programming language that works in Microsoft Excel. This program streamlined the process by first asking the user what year that they would like to look at, and then examining all of the data and returning the total Daily volume traded and the total return percentage over that year. The process was further made more user friendly by the use of a button, and a message box asking for the year that is to be examined. The process is finished with a pop-up window stating how long it took the computer to run this code.

## Results
### Initial Coding to Begin the Analysis

The initial step was to Create and Initialize an array of all of the different tickers in the data, and to represent the data as words (As String). This told the program what the subsets of data that it was going to look at.
```
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
```

The next step was to tell the progarm to examine all of the rows in the data sheet
`RowCount = Cells(Rows.Count, "A").End(xlUp).Row`

To set the initial value of the ticker to 0 
`tickerindex = 0`

Then to set the values of the Volumes, Starting Prices and Ending Prices to 0, and then to do it for each ticker (a "for" statement)

```
For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
```
These initial steps set the program up to start at the beginning of each different sets of tickers, preparing it to begin the actual analysis.

### The `for` loops to begin looking at the information in each individual ticker.
The next set of steps was to tell the program to look at each individual ticker.

The first code was to have the program loop through every row of information in the spreadsheet. 
`For i = 2 To RowCount`

The second line of code was to tell the program to begin to add the Volumes of trades on every row that it examined.
`tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value`

The third line of code was to check if the row that is being examined was the first row of that ticker. This code translates as: if the current row is equal to the current ticker and the row above it is not equal to the ticker, then it is the first row of the ticker, and to begin tracking the data.
```
If Cells(i, 1).Value = tickers(tickerindex) And Cells(i - 1, 1).Value <> tickers(tickerindex) Then
            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
        End If
```

The fourth line of code was to check if the row that is being examined is the last row of the ticker, and if it is, to move on to the next ticker. This code translates as; if the current row is equal to the current ticker and the row below it is not equal to the ticker, then it is the last row of the ticker and to stop tracking the data. If this statement is true, move on to the next ticker in the sequence.
```
If Cells(i, 1).Value = tickers(tickerindex) And Cells(i + 1, 1).Value <> tickers(tickerindex) Then
            tickerEndingPrices(tickerindex) = Cells(i, 6).Value
	    tickerindex = tickerindex + 1
        End If
```

### The `for` loop to apply the analysis to all tickers in the data, and to output them into the correct location in the spreadsheet.
This code was to tell the program to take the above `for` loop and apply it to each ticker in sequence. Then to take that data and place it in the correct cell in the All Stocks Analysis sheet
```
For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
```

### Formatting the data to make the analysis easier to read at a glance
This step was to format the cells in the table to make them easier to read.
```
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("a1").Font.Bold = True
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    Columns("C").AutoFit
```
This code is telling the sheet to look a certian way, in this order;
- Cells A3, B3, and C3 are bold
- The bottom edge of A3, B3, and C3 will have a thick bold line
- Cell A1 will be bold
- The data in column B will have the format $0,000,000.00
- The data in column C will have the format 0.0%
- The sizing of the cells in columns B and C will format themselves automatically to the correct size

These lines of code are telling the program to change the color of the cells (in column C, between rows 4 and 15) to green if they are returning a positive precentage, or red if they are returning a negitive percentage.
```
dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
```
![2017](https://github.com/chefcramer/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.png)
![2018](https://github.com/chefcramer/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.png)


## Summary
There are many advantages to refactoring code to make it run much smoother and quicker. It is the difference between publishing a first draft and publishing a polished and edited final draft, its much easier to read (execute) and makes much more sence. Refactoring this code made the program run .22 seconds faster on my machine from the original code to this refactored code. It seems like such a small difference of time but it is almost 30% faster, which makes a HUGE difference when your data set is much much larger, 10 or 100 thousands instead of just over 3000. The advantages far outweigh the disadvantages in refactoring code. The disadvantages i ran into are keeping the values straight in your head, i ran into several spelling mistakes in my code that (tickersindex instead of tickerindex) that really messed me up and was almost impossible to spot until i went line by line to find it.

It seems that the original code is SLIGHTLY shorter than the refactored code, but they accomplish the same thing (within this data set). The refactored code does not use a nested `for` loop, and the original does. The refactored code is much more strightforward in this aspect, as I have encountered some issues keeping my variables straight in a nested loop. The refactored code can also be scaled to any data set, the original code is much more limited. The original code can be scaled as well, but it would require more work to add more data to examine, while the refactored code is ready to scale now.