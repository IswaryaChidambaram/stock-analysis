This analysis shows the total daily volume calculated for all the tickers and the return is calculated based on the starting price and the ending price for the year 2017 and year 2018. The code is refactored so that to expand the dataset to include the entire stock market over the last few years in less amount of time.
## Results of 2017 Analysis:
As a part of Refactoring tickerIndex is defined to access the ticker, tickeVolumes,tickerStartingPrices and tickerEndingPrices. 
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Worksheets(yearValue).Cells(i, 8).Value
tickerVolumes(tickerIndex) stores the total volume of the selected ticker.
If Worksheets(yearValue).Cells(i, 1).Value = currentTicker And Worksheets(yearValue).Cells(i - 1, 1).Value <> currentTicker Then
          tickerStartingPrices(tickerIndex) = Worksheets(yearValue).Cells(i, 6).Value
        End If
Here tickerStartingPrices(tickerIndex) stores the starting prices of the selected ticker provided Cells(i, 1).Value is the first row of the selected ticker and Cells(i - 1, 1).Value is not the  currentTicker
 If Worksheets(yearValue).Cells(i, 1).Value = currentTicker And Worksheets(yearValue).Cells(i + 1, 1).Value <> currentTicker Then
            tickerEndingPrices(tickerIndex) = Worksheets(yearValue).Cells(i, 6).Value
            tickerIndex = tickerIndex + 1
        End If
         
Here tickerEndingPrices(tickerIndex) stores the ending prices of the selected ticker provided Cells(i, 1).Value is the last row of the selected ticker and Cells(i -+1, 1).Value is not the  currentTicker. If the condition satisfies, tickerIndex is incremented to 1.

### Stock Performance of 2017:
The Returns of the all the tickers except TERP has a positive return. DQ has the highest positive return of 199.4% 

### Screenshot of 2017 Stock Performance:
<img width="1440" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/83719443/124291764-e4dd2000-db22-11eb-9f99-9c3dff794b05.png">

### Stock Performance of 2018
Only the tickers ENPH and RUN has a positive return. All the other tickers has a negative return being JKS with the negative return of -60.5%

### Screenshot of 2018 Stock Performance

The wanted columns have to be plotted for the chart. Otherwise the chart will be formed for the entire table which leads to inaccurate values.

### Screenshot of the Analysis based on Goal:
![Outcoms based on Goals](Resources/Goals%20screenshot.png)

## Results:

### Based on Launch date:
* We can see that theater campaigns have been the most successful in May and less successful in December. Theater Campaigns have failed the most in december. Intrestingly,    campaigns generally dont get cancelled in October. 
* The summer season looks favourable for the theater campaigns.

### Based on Goal amount:
* We can see that when the goal amount<$1000 , the "plays" campaign tends to be the most successful. When the goal amount is more than $40000, the "plays" campaign tends to fail.

### Recommendations:
The dataset doesnt provide the information about what age group  were involved in the different category of campaigns. Having the age group information would be helpful to target better in the setting of goals and the amount received. That will also help to determine which campaign will work in different months.
