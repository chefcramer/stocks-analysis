# VBA of Wallstreet

## Overview of Project
	This project is to analyze the overall performance of a data set of 12 different "Green" companies to determine which is the best investment opportunity. The Data was organized, analyzed and formatted using VBA (Visual Basic), a programming language that works in Microsoft Excel. This program streamlined the process by first asking the user what year that they would like to look at, and then examining all of the data and returning the total Daily volume traded and the total return percentage over that year. The process was further made more user friendly by the use of a button, and a message box asking for the year that is to be examined. The process is finished with a pop-up window stating how long it took the computer to run this code.

## Results
### Initial Coding to Begin the Analysis\

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

## Summary
