# VBA Challenge

## Overview of Project

Steve has recently graduated with a degree in finance. His parents want to invest and they are passionate about renewable energy. They want to make all their investments in a renewable energy firm called DAQO energy which makes Solar panels. Steve wants to look into the stock options of DAQO and he is also concerned about diversifying their funds and investing in different Green Energy Stocks.
He wants to analyze a handful of green energy stocks along with DAQO stocks. He has created an excel file with the required stock data of the companies he wants to analyze. This sheet contain the company ticker, date, opening and closing stock price, volumes of the particular day.
Based on the available data, we have to create a solution which provides Steve with the total volume and returns from each company so that he can make an informed decision on where to invest his parent's funds.


### Purpose

The purpose of the project is to provide Steve with the total volume and return for each green energy company in his sheet so that he is able to make the most feasible decision to invest his funds

## Analysis and Challenges

The results for this project has been generated using VBA.
There are 12 companies in this sheet, and each of these companies have data for several date belonging to two years - 2017 and 2018.
The aim is to find the total volume and the return for each company for the years 2017 and 2018.
The total volume has been calculated by using a loop for all the records in the sheet:

For i = 2 To RowCount
If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If

Also the starting and ending prices for each company has been stored so that the return can be calculated for each company

Below is the code to calculate the starting price, when the current row is the first row of a company

If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If

Below is the code to calculate the ending price. It is calculated when the current row is the last row of a company

If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

After calculations, the output values are populated in the target sheet

For i = 0 To 11

Company name and company volumes are populated using the below code
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       
Below is the code to calculate the return
       Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        Worksheets("All Stocks Analysis").Activate
        
    ' Close loop i
    Next i

The challenges have been:
1. Creating the VBA code with multiple loops
2. The volume of the data should be increased to make the results better


## Results

The stock results from 2017 is in the below screenshot:
![Screen Shot 2021-05-15 at 8.04.23 PM](https://i.imgur.com/JLcZDXX.png)


The stock results from 2018 is in the below screen shot:
![Screen Shot 2021-05-15 at 8.05.25 PM](https://i.imgur.com/Z5Gk5OY.png)

The comparison of the 2017 vs 2018 is as below:
![Screen Shot 2021-05-15 at 8.19.37 PM](https://i.imgur.com/aJeyr9P.png)

It can be seen that only two companies have managed to have positive returns in 2018.
The returns of DAQO has dropped to -62.6%
The two best companies to invest in are as below:
1. RUN - returns have increased from 5.5% in 2017 to 84.0% in 2018
2. ENPH - returns have dropped from 129.5% in 2017 to 81.9% in 2018. However it is largely profitable than the other companies

The time required to execute the 2017 process using the old code is as below:
![Screen Shot 2021-05-15 at 8.25.23 PM](https://i.imgur.com/S28uoej.png)


The time required to execute the 2017 process using the refactored code is as below:
![Screen Shot 2021-05-15 at 8.26.45 PM](https://i.imgur.com/Uhu7rBH.png)

The timings have improved after using the refactored code.

The time required to execute the 2018 process using the old code is as below:
![Screen Shot 2021-05-15 at 8.28.07 PM](https://i.imgur.com/uKsU06N.png)

The time required to execute the 2018 process using the refactored code is as below:
![Screen Shot 2021-05-15 at 8.30.35 PM](https://i.imgur.com/uS783Q5.png)

The timings have improved after using the refactored code.

## Summary

a. What are the advantages or disadvantages of refactoring code?

The advantages of refactoring:
1. It improves the design of the code. It provides a more generic way of implementing the code.
2. It makes the code faster
3. It provides a fresh perspective to the code and helps to increase its maintainability

The disadvantages of refactoring:
1. It is time consuming
2. While the solution may become more generic, the complexity may increase as well
3. It is more difficult to debug


b. How do these pros and cons apply to refactoring the original VBA script?

The pros:
1. The code has now become more generic
2. The processing times have become faster

The cons:
1. The introduction of arrays for startprice, endprice, tickervolumes has increased the complexity of the code
2. It has become more difficult to debug due to the use of arrays and loops while populating the target fields




