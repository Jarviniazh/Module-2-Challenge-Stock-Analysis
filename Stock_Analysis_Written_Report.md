# Green Stock Analysis with VBA

## Overview of Project

### Purpose
Though we assisted Steve to analyze a handful green energy stocks for his parents by using VBA code to automate the analyses previously, Steve is worried about if he put more stocks in the future the current Marco may not perform as perfect as right now. Hence, we would refactor the original code by means of improving the logic, to enhance its efficiency and make sure it is capable to reuse with even larger stock market data in the future.  

## Results
In order to improve efficiency of the script, we went over the original one and noticed we could make the nest loop fresher. Therefore, one index variable, tickerIndex and three arrays (tickerVolumes, tickerStartingPrices, and tickerEndingPrices) were added. The new nested loop should only execute once the corresponding index is detected. The below is the refactored VBA script:

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
    
    '1a) Create a ticker Index and set it equal to zero before looping
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
        tickerVolumes(tickerIndex) = 0
    
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            If Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            

            '3d Increase the tickerIndex.If the next row's ticker dosen't match the previous one
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
               
        'End If
    
        Next i
    
    Next tickerIndex
    
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

#### The output for 2017 and 2018
- 2017

    ![image](https://github.com/Jarviniazh/Module-2-Challenge-Stock-Analysis/blob/main/Resources/Outputs%20of%202017.png)

- 2018
    
    ![image](https://github.com/Jarviniazh/Module-2-Challenge-Stock-Analysis/blob/main/Resources/Outputs%20of%202018.png)

#### Execute time of original code and refactoring code
- Original code running time

    ![image](https://github.com/Jarviniazh/Module-2-Challenge-Stock-Analysis/blob/main/Resources/VBA_Challenge_2017_Original.png)
    ![image](https://github.com/Jarviniazh/Module-2-Challenge-Stock-Analysis/blob/main/Resources/VBA_Challenge_2018_Original.png)

- Refactoring code running time

    ![image](https://github.com/Jarviniazh/Module-2-Challenge-Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)
    ![image](https://github.com/Jarviniazh/Module-2-Challenge-Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary
#### What are the advantages or disadvantages of refactoring code?
- Pro: Refactoring code makes it easier to read and maintain. In addition, refactoring will speed up the performance of the code and save memory. Besides, people can target a specific section of code other than rewriting, so they do not need to do too much to improve small portions of a code. 

- Cons: The process of refactoring is very time-consuming. And in case if it went wrong, we would have to waste much more time in solving the problem and there are probable chances that it may go wrong due to complexity of the code. 

#### How do these pros and cons apply to refactoring the original VBA script?
- Pro: The refactoring script save way more execute time than the original one. And since we give every variable a dimension such as tickerIndex, the well-defined variables make the script easier to read even the user change in the future. Meanwhile, based on Steve’s request, he could expand the dataset to included the entire stock market over the last few years with reasonable execute time later.

- Cons: To make the script more efficient, we introduce more variables, which are needed to declare them properly. We must have a clear understanding of each variable and the role it will play in the new script, especially the data type. For example, in this project if we dim totalVolunm as integer, there will be a bug report pop up

