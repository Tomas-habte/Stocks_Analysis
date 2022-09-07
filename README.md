# An Analysis of Stocks using Excel VBA
## Project Overview
---
The purpose of this analysis is to refactor code which is analyzing the performance of several stocks. By refactoring, we are going to see if we can make the VBA script run faster in order to allow us to analyze the entire stock market in a timely manner.  

---

## Results

**Stock Performance in 2018 vs 2017

As shown below, using VBA we were able to run a script that pulled stock data from two different sheets and calculated the returns. We also used conditional  formatting to illustrate the performance of each stock. This showed us that 2017 was a very good year with high returns for all but one ticker (TERP). It also showed us that the stocks performed very poorly in 2018 with the two exceptions being ENPH and RUN. 

![2018 Stocks Performance ](https://user-images.githubusercontent.com/111667387/188781000-56f10257-7f7d-456e-b045-d4f6ca040aeb.png)
![2017 Stocks Performance ](https://user-images.githubusercontent.com/111667387/188781204-5eec2113-682e-4184-9c7e-2a1af9cd3b90.png)

**Code Refactoring 

Intially, our code did not have a ticker index but we used that to refactor the code as shown below.

'1a) Create a ticker Index
   tickerIndex = 0
   
   '1b) Create three output arrays
   Dim tickerVolumes(12) As Long
   Dim tickerStartingPrices(12) As Single
   Dim tickerEndingPrices(12) As Single
   
   ''2a) Create a for loop to initialize the tickerVolumes to zero.
   For i = 0 To 11
       tickerVolumes(i) = 0
   Next i
   
   '2b) Loop over all the rows in the spreadsheet.
   For i = 2 To RowCount
   
       '3a) Increase volume for current ticker
       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

Compared to the intital code we used below:

  '3a) Initialize variables for starting price and ending price
   
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    '3b) Activate data worksheet
    
    Worksheets(yearValue).Activate


    '3c) Get the number of rows to loop over
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '4) Loop through tickers
    
     For i = 0 To 11
     ticker = tickers(i)
     totalVolume = 0
     
          '5) loop through rows in the data
          
          Worksheets(yearValue).Activate
           
           For j = 2 To RowCount
                '5a) Find total volume for current ticker
                If Cells(j, 1).Value = ticker Then

                     totalVolume = totalVolume + Cells(j, 8).Value
                End If
      
**Code Performance

After refactoring the code, we created a timer in VBA to compare the difference performance between the old and new code. As shown below, refactoring the code was able to make the script run signficantly faster. 
Old ![2017 old performance](https://user-images.githubusercontent.com/111667387/188785126-04e56f4d-8893-40ec-a948-ae27123cc9bb.png) vs New ![2017 new performance](https://user-images.githubusercontent.com/111667387/188785135-4e863359-2e1e-4ef0-bc67-e186f094ff54.png)
Old ![2018 old performance](https://user-images.githubusercontent.com/111667387/188785180-e0e2c89b-146a-4991-9819-55876f9efaa4.png) vs New ![2018 Stocks Performance ](https://user-images.githubusercontent.com/111667387/188785198-5ed2d5e4-c636-434f-9cbc-5147ca35cc7c.png)

 
