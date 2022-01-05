# Refactoring VBA to Improve Performance of Stock Analysis

## Overview of Project

### Purpose

Steve (client) has asked that the analysis previously performed on a single stock, DAQO, be upscaled to analyze the entire data set. This will require that the initial code be refactored to run more efficiently and over a larger data set.  The purpose of this project is to deliver that refactored code and provide Steve with an analysis of any evidence of improved efficiency in code run time.

### Background 

The initial code delivered to Steve looked at 12 green-energy tickers by looping through each row of data 12 times (one time for each ticker) to extract information on daily volume and return.  While this technique provided insight on the current dataset in a matter of seconds, it is not scalable to a larger dataset.  

## Results

2017 was a better year for green energy stocks than 2018.  (See images below.) Of the 12 green energy tickers analyzed, 11 had a positive return in 2017 where only 2 had positive returns in 2018.  Steve’s green-passionate parents currently only hold DAQO stock, which had a negative return in 2018 despite strong performance in 2017.  Steve should advise that his parents diversify to various types technology within the green energy sector, but also into other green industries as well.  It should be noted that two years of data is not enough to observe trends. 

Image: 2017 and 2018 POST-refactored data and performance times

<img width="484" alt="VBA_Challenge_2017_Post-Refactored_Timer" src="https://user-images.githubusercontent.com/93740725/148283746-82a54c95-06e6-420a-9be4-f25d57037065.png">
<img width="468" alt="VBA_Challenge_2018_Post-Refactored_Timer" src="https://user-images.githubusercontent.com/93740725/148283754-dea222ec-167f-465f-8376-2eb9f5c5d8ee.png">


The good news is, the refactored code will help Steve identify other green industries with consistent YOY performance because the new code significantly cut down on run time!

Image: Runtime Summary Table

<img width="251" alt="VBA_Challenge_Run_Time_Summary_Table" src="https://user-images.githubusercontent.com/93740725/148283817-58f2679a-a7f7-4e9a-ba2a-993bb11a8ae6.png">

Instead of repeating the for loop 12x, all data was collected into output arrays in one pass.  This resulted in 6 to 7 times faster performance. 

### Code Breakdown

The set up code stayed the same:

	Sub AllStocksAnalysisRefactored()
    
    'make variables for collecting run time
        Dim startTime As Single
        Dim endTime  As Single

    'collect year and start time of timer
        yearValue = InputBox("What year would you like to run the analysis on?")
        startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
        'create title and year
            Worksheets("AllStocksAnalysis").Activate
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

Next was to define an index variable for the arrays, make output arrays, and set initial volumes in the array to zero:

    '1a) Create a ticker Index and set to zero
        tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
        Next i
In the new code, we only needed one pass through the rows to collect data from all 12 tickers into each of the three output arrays. The tickerIndex variable made this possible. Also, we dropped the AND conditionals from the original code because we are looking at all tickers instead of just one:

    '2b) Loop over all the rows in the spreadsheet
        RowCount = Cells(Rows.Count, "A").End(xlUp).row
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                   tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
        '3c) check if the current row is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
        '3d) If the next row's ticker doesn't match, increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            
            End If
        Next i

Finally, -   the output is changed to pull from arrays instead of one variable:

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        Worksheets("AllStocksAnalysis").Activate
        
        For i = 0 To 11
            Cells(i + 4, 1).Value = tickers(i)
            Cells(i + 4, 2).Value = tickerVolumes(i)
            Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        Next i
   
Formatting is similar to pre-refactored code: 

    'Formatting
        Worksheets("AllStocksAnalysis").Activate
        
        'title format
        With Range("A1")
            .Font.Bold = True
            .Font.Size = 14
        End With
        
        'headers format
        With Range("A3:C3")
            .Font.Bold = True
            .Font.Size = 13
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        
        'numbers format
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        
        'columns width format
        Columns("B").AutoFit
        Columns("C").AutoFit

        dataRowStart = 4
        dataRowEnd = 15
        
        'Conditional formatting to highlight performance
        For i = dataRowStart To dataRowEnd

            If Cells(i, 3) > 0 Then
                'Color the cell green
                Cells(i, 3).Interior.Color = vbGreen

            ElseIf Cells(i, 3) < 0 Then
                'Color the cell red
                Cells(i, 3).Interior.Color = vbRed
            
            Else
                'Clear the cell color
                Cells(i, 3).Interior.Color = xlNone

            End If

        Next i
       
Data is collected on the run time the same way as the pre-refactored code, but a summary table is added to collect run times for each year in each version of code:

    'collect end time of run
        endTime = Timer
    
    'display total run time in message box
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    'add timer data to summary table
        Worksheets("RefactorAnalysis").Activate
    
        If yearValue = "2017" Then
            Cells(2, 3).Value = (endTime - startTime)
        Else
            Cells(3, 3).Value = (endTime - startTime)
        End If
        
    'format summary table
        Range("B1:D1").Font.Bold = True
        Range("A1:A3").Font.Bold = True
        Range("b2:d3").NumberFormat = "0.00"
        
	End Sub

## Summary

### What are the advantages or disadvantages of refactoring code?

Refactoring code can be advantageous because it helps legacy code run more efficiently or read more clearly for future users.  This cuts down on run time, frees up computer memory, and potentially allows the code to run on larger and larger data sets.   One drawback is that as you update code, you’ll potentially break some code and have to re-test and re-debug code that was already previously working. A second drawback could be that the person doing the refactoring might not be the original author, so the original author’s intent may not be obvious unless the code comments are adequately detailed. 

### How do these pros and cons apply to refactoring the original VBA script?

The legacy code in this project did adequate work but did not run quickly.  One clear advantage of refactoring the green-stock analysis code was that the new run time was 6 times faster than the old code. Too, the new code could easily be used on anther data set if Steve wanted, all he would need to do is manually input the tickers into the tickers array. 

One major disadvantage of the process of refactoring this code was that it took a while to debug.  A main hang-up, for me, was the old code used IF AND conditionals, and the new code does not.  Since the original code was only looking at one ticker at a time, it the IF ANDs made sense.  “IF the current ticker = DQ AND the next does not, then you have found the last entry.”  I erroneously copied and passed my old IF AND code in, but it broke the new code.  This is because each row was not just being evaluated if it had DQ data, it was being evaluated for any ticker data; the AND was invalid.  

Using arrays to store information and create output made the refactored code run faster and visibly look cleaner.  A disadvantage, though, is that arrays can be harder for new programmer to understand.  Looking for one ticker, DQ, in each row and then pulling info from just those rows is intuitive to most people: it is how many folks would get the data manually if they were reading a table.  	

A final short coming of the refactored code is that it requires the end user to manually input the tickers into an array.  If this code were to be further refactored, VBA could be used to [identify unique tickers](https://stackoverflow.com/questions/36044556/quicker-way-to-get-all-unique-values-of-a-column-in-vba)  in column A, count and add those to an array, and output them to the analysis sheet. 
