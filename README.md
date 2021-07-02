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
![2017 Stock Performance Analysis]()

## Analysis based on Goals:
Baesd on the goal amount, number of campaigns which have been successful, failed, cancelled have been calculated. The Total projectsis the sum of the successful, failed, cancelled campaigns.Also the % of all these outcomes has been calculated.COUNTIFS is used to here to calculate the number of campaigns for all the goal amounts in all the outcomes in the"plays" subcategory. A line chart has been created which covers the goals, % successful, % failed, % canceled. We can see that when the goal amount<1000$ , the campaign tends to be the most successful. When the goal amount is more than 40000$, the campaign tends to fail.

### Challenges which could be encountered:

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
