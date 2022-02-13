# Green Energy Stock Analysis

## Overview of Project

>Our friend Steve is researching Green Energy stocks for his parents to invest in and has asked for out assistance in helping him analyze the data for 12 different green stocks in 2017 and 2018.   


## Purpose
> The purpose of this project is to compare the performance of the green stocks between 2017 and 2018 and then to refactor the original code to make it more efficient by decreasing the time it takes to run.

---   



## Results

### Stock Performance

>In reviewing the green stocks data for Steve, we can see that the returns in 2017 were significantly better than the returns in 2018.  For 2017, with the exception TERP, all of the stocks had a positive return on investment with several increasing over 100%.  Unfortunately, for 2018 that positive trend did not continue as all except two of the green stocks experienced a negative return on investment.  The two stocks that did have a positive return were RUN and ENPH and only RUN out performed its 2017 return by going from 5.5% to 84%.  Even though the 2018 stocks took a negative hit for the majority of them it wasn't enough to offset all the positive gains from 2017.
 

![](https://github.com/timbialek/Stock_Analysis/blob/main/Resources/All_Stocks_Returns_2017.PNG) ![](https://github.com/timbialek/Stock_Analysis/blob/main/Resources/All_Stocks_Returns_2018.PNG)

### The Coding

>The goal for the refactored coding is to loop through the data one time and collect all of the information. Below I will discuss the key parts of the refactored code but if you would like to see all of the coding here is a link to the  [Original Code](https://github.com/timbialek/Stock_Analysis/blob/main/Resources/All_Stocks_Analysis_Orignal.txt), the [Refactored Code](https://github.com/timbialek/Stock_Analysis/blob/main/Resources/All_Stocks_Analysis_Refactored.txt) and the excel file where the code can be found and run for the
[Stock Analysis](https://github.com/timbialek/Stock_Analysis/blob/main/VBA_Challenge.xlsm).  In the xlsm file the AllStocksAnalysis module contains the original code and AllStocksAnalysisRefactored is the updated code.  To easily test the performance between the original and refactored code there are buttons on the All Stocks Analysis tab to run each module.

First, we are going to set up a new variable called tickerIndex and three arrays that will store out output.  This is important since it allows us to store the data as we loop through it.


    '1a) Create a ticker Index
     tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For tickerIndex = 0 To 11
       tickerVolumes(tickerIndex) = 0

The next steps in the coding involve setting the new tickerIndex variable to zero and then using it as the index when calculating the tickerVolumes, then in the if-then statements to get the tickerStartingPricces and tickerEndingPrices and finally we use it in a script that increases the tickerIndex. 
     
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        
                
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                    
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rows ticker doesn't match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6)
                    'tickerIndex = tickerIndex + 1 <-- seems to run a little faster when put here
         End If
         

        '3d Increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                   tickerIndex = tickerIndex + 1
            
            
        End If
    
    Next i
    
    Next tickerIndex


For step 4, we use the tickerIndex to output the data we collected in the arrays for the Ticker, Total Daily Volume and the Return columns in the excel file


 '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
  
    For tickerIndex = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + tickerIndex, 1).Value = tickers(tickerIndex)
        Cells(4 + tickerIndex, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + tickerIndex, 3).Value = (tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex)) - 1
        
        
    Next tickerIndex

The rest of the coding is related to the formatting of the final output and a pop-up box with the time it takes the script to run.  Neither of these sections were refactored during this project but if you would like to see the code is at the bottom of the [Original Code](https://github.com/timbialek/Stock_Analysis/blob/main/Resources/All_Stocks_Analysis_Orignal.txt) and [Refactored Code](https://github.com/timbialek/Stock_Analysis/blob/main/Resources/All_Stocks_Analysis_Refactored.txt) code documents.

---




### Analysis of the Refactored Script Execution Times

In comparing the run times of the refactored script for 2017 and 2018 to the original script times it can be seen that the refactored script had a significant improvement in going from over .7 seconds to now just over .1 second.  This is an 84% improvement for 2017 and 83% improvement for 2018. 

Refactored Script Times:

![](https://github.com/timbialek/Stock_Analysis/blob/main/Resources/Refactored_Code_2017_time_stamp.PNG) ![](https://github.com/timbialek/Stock_Analysis/blob/main/Resources/Refactored_Code_2018_time_stamp.PNG)

Original Script Times:

![](https://github.com/timbialek/Stock_Analysis/blob/main/Resources/Original_Code_2017_time_stamp.PNG) ![](https://github.com/timbialek/Stock_Analysis/blob/main/Resources/Original_Code_2018_time_stamp.PNG)



>There are both advantages and disadvantages to refactoring code.  The main advantage of refactoring code is that it allows you to make the code more efficient.  This can include taking fewer steps to accomplish a task, using less memory, and improving the logic of the code to make it easier for future users to read.  Often the first attempt at written code is not always the most efficient so refactoring can help to make the code better.  Some disadvantages of refactoring code is the times it to perform the refactoring of the new code and that new code could introduce bugs into the program which would then need to be fixed.  


### Pros and Cons of refactoring the original VBA script
>When it comes to the refactoring of the original VBA script the refactored code enhanced the run time.  It shaved about 6 seconds off the run time which is a significant improvement.  While figuring out how to refactor the code there were several instances where I made errors that created bugs in the program for which I then had to spend additional to correct.   In the scope of this project which only reads a few thousand rows of data this might not have been an efficient use of time but if we were to apply this code to few million lines of data it would have been very worthwhile time investment.


## Resources:
* Data Source: green_stocks.xlsx, VBA_Challenge.vbs
* Software: Excel, VBA	

