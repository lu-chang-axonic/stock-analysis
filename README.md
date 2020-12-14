
<h1 align="center">Stock Analysis</h1>

## Overview of Project
The project refactors the module 2 solution code by using functions such as For, IF, Array, and Loop. The codes efficiently run through all stocks to calculate the trading volume and return by ticker. It also is user friendly by having input message box and runtime message to help the user understand what was done with how much time. 

## Results

#### Stock Performance Comparison Between 2017 and 2018

The two tables below show the outcome of the analysis: 

![](https://github.com/lu-chang-axonic/stock-analysis/blob/main/images/2017%20Result%20for%20Stock%20Analysis.png)
![](https://github.com/lu-chang-axonic/stock-analysis/blob/main/images/2018%20Result%20for%20Stock%20Analysis.png)

The color coded result made it easy to identify that 2017 was a better year of performance for the stocks under discussion. 

#### Execution Time Comparison
The two message boxes below how the run time of the analysis was before the refactoring, saved in fiel "VBA_Challenge" :
![](https://github.com/lu-chang-axonic/stock-analysis/blob/main/images/VBA_Challenge_2017.PNG)
![](https://github.com/lu-chang-axonic/stock-analysis/blob/main/images/VBA_Challenge_2018.PNG)

The two message boxes below how the run time of the analysis was after the refactoring by using an Index, saved in "VBA_Challenge Using Index:
![](https://github.com/lu-chang-axonic/stock-analysis/blob/main/images/Enhanced%20Run%20Time%202017.PNG)
![](https://github.com/lu-chang-axonic/stock-analysis/blob/main/images/Enhanced%20Run%20Time%202018.PNG)

So, the performance of the code was improved after the refactoring, as reflected by the reduced execution time. The array and nested loop has made the collection of information very efficient (shown below). 

For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'If Cells(i, 1).Value = tickers(tickerIndex) Then
        'TotalVolume = TotalVolume + Cells(i, 8).Value
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         tickerIndex = tickerIndex + 1
         End If
          
            '3d Increase the tickerIndex.
         
                 
    Next i


## Summary
#### What are the advantages or disadvantages of refactoring code?
###### Advantages
1. It is a very efficient way to reuse code, so that one does not need to write code line by line. 
2. The previous code is proven to be working. This saves debug time too.
3. It is great way to learn for beginner programmers. 

###### Disadvantages
1. Without writing the code from scratch, the programmer could have easily missed nuances that could cause problems down the road.
2. If the new problem is not identical to the old problem, the refactoring could ended up consuming more time than writing from scratch.
3. It is an easy way to learn, but the learned knowledge may not be solidly yours because everything is already pre-written.

#### How do these pros and cons apply to refactoring the original VBA script?
In doing this challenge, I found myself following the question and the original VBA scripts really easily and finished the project very quickly. However, I am not sure if being given a different problem, I would be able to write the script without referring back the sample repeatedly as I do not remember the codes in details.

