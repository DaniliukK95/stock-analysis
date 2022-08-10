# stock-analysis
Repo for Module 2


beginning of challenge 2
# Stock Analysis
Challenge 2 Assignment 


## Overview

### Background
Steve recently graduated with a finance degree! His parents are passionate about green energy and decided to invest in the stock ticker “DQ”. Steve, being the caring son that he is, wants to look further into their investment along with other potential investments that can fit their portfolio. Steve has come to us to analyze a handful of energy stocks including “DQ”. In Module two we wrote code in VBA to automate analysis for any stock in the data set. By the end of Module two, we created a VBA script that analyzes the information for the years 2017 and 2018 of the twelve stocks we have in the data set. Although great for our data set, what if we wanted to do the entire stock market over more than 2 years? The code we have already created might not be sufficient.
### Purpose
The purpose of this challenge is to edit/refactor our original code to take fewer steps overall, use less memory, increase code performance, and improve the logic of the code for digestible reading. This will make the code more efficient in many areas. We will analyze the performance of the stocks to each other, and between the 2 years. We will also compare our old VBA script to the new one based on run time after refactoring.

## Results of our Analysis

### Comparison of Stock Performance Between 2017 and 2018 
By looking at tables from the analysis, it’s clear to see that green energy stocks performed better in 2017 than they did in 2018. It seems that it was a very bullish year for most green energy stocks. We can see the tables here:

![2017_stocks.png](https://github.com/DaniliukK95/stock-analysis/blob/main/Resources/2017_Stocks.png) ![2018_stocks.png](https://github.com/DaniliukK95/stock-analysis/blob/main/Resources/2018_Stocks.png)

We used formatting to color code which stocks went up in price from the beginning of the year to the end of the year. This is seen as green for increase in price and red for a decrease in price. This was achieved with the `interior` function connected to the `color` function. By using an If-Then statement combined with an If-Else statement, we were able to automate VBA to distinguish which stocks had a positive or negative price change, and then to color code them accordingly. 
```
If Cells(i, 3) > 0 Then
Cells(i, 3).Interior.Color = vbGreen
Else
Cells(i, 3).Interior.Color = vbRed
End If
```
This makes the analysis easier to look at from a glimpse! We can clearly see majority of the stocks are green in 2017 compared to 2018. Looking at the total daily volume to me doesn’t really provide much information. The reason why I say this is because if a stock is valued at $1.00, it does not take much to buy many shares. On the other hand, if a stock is worth $1000 per share, it would not be surprising to see a lower daily volume than the cheaper stock. What would be more tactical to use regarding volume is comparing It to its valuation. We can also see that there isn’t a clear answer as to if the total daily volume between the years had some type of correlation to its return. Although the tickers “ENPH” and “RUN” continued to run bullish in 2018 with increasing volume, “TERP” had an increase in volume but stayed bearish in both years. Some of the other stocks increased in volume but instead resulted in negative returns. So not much can be determined by this data. From the color coding and percentages, it is clear to see that the beginning of 2017 was the year to invest. Since Steve’s parents invested, I’m assuming after 2018, this does not mean that they made a bad investment! The stock market has many variables to factor in like news, economy, and even supply and demand from a price action viewpoint. This bearish 2018 could just be a correction in price for the rally it had in 2017 and could possibly make another run in 2019. Or maybe there was a green energy bubble in 2017 which popped and is now making its way down. To most people, this probably is a terrible sign, but to another investor this could be a discounted price! 

### Execution Time of the Original Script VS the Refactored Script
In my process of refactoring the VBA script, I noticed extreme differences immediately. I also included suggestions on speeding up run time in VBA using other code from research. The run time for the original code for 2017 and 2018 was 0.67 seconds and 0.62 seconds respectively as seen in these screen shots: 

![original_VBA_runtime_2017.png](https://github.com/DaniliukK95/stock-analysis/blob/main/Resources/Original_VBA_RunTime_2017.png) ![original_VBA_runtime_2018.png](https://github.com/DaniliukK95/stock-analysis/blob/main/Resources/Original_VBA_RunTime_2018.png) 

After refactoring for the first time, I was able to achieve an incredible speed of 0.097 seconds for 2017 and 2018. 

![VBA_Challenge_2017_before internet help.png](https://github.com/DaniliukK95/stock-analysis/blob/main/Resources/VBA_Challenge_2017_before%20internet%20help.png) ![VBA_Challenge_2018_before internet help.png](https://github.com/DaniliukK95/stock-analysis/blob/main/Resources/VBA_Challenge_2018_before%20internet%20help.png)

That’s a whole 0.6 seconds decreased, which doesn’t seem like much, but if we had hundreds of thousands of data points, it could make a difference. After adding 4 extra lines of code from [Tips to Speed up VBA Code](https://eident.co.uk/2016/03/top-ten-tips-to-speed-up-your-vba-code/) I was able to further drop the run time to 0.062 seconds for both 2017 and 2018. 

![VBA_Challenge_2017.png](https://github.com/DaniliukK95/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png) ![VBA_Challenge_2018.png](https://github.com/DaniliukK95/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

One of the main changes that helped bring down the run time was to get rid of the nested for loop and to use one condition only in the If-Then statements. The nested for loop to get the total daily volume and the return values: 
```
For i = 0 To 11
ticker = Tickers(i)
totalvolume = 0
Sheets(yearvalue).Activate
For j = 2 To RowCount
If Cells(j, 1).Value = ticker Then
totalvolume = totalvolume + Cells(j, 8).Value
End If
If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
startingprice = Cells(j, 6).Value
End If
If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
endingprice = Cells(j, 6).Value
End If
Next j
```
Was converted into two For loops with one condition in each If-Then statement:
```
For i = 0 To 
TickerVolumes(i) = 0
Next i
For i = 2 To RowCount
TickerVolumes(tickerindex) = TickerVolumes(tickerindex) + Cells(i, 8).Value
If Cells(i - 1, 1).Value <> Tickers(tickerindex) Then
TickerStartingPrices(tickerindex) = Cells(i, 6).Value
End If
If Cells(i + 1, 1).Value <> Tickers(tickerindex) Then
TickerEndingPrices(tickerindex) = Cells(i, 6).Value
tickerindex = tickerindex + 1
End If
Next i
```

When working with a classmate, they had a nested For loop in their code which ended up taking just as much time as the original code. When comparing my time to his, it was clear that the nested For loop was doing more harm than good. The code used from the internet was `Application.ScreenUpdating = False` and `Application.Calculation = xlCalculationManual`. These help to prevent the screen from flickering and updating while running the code, and for the application to prevent calculations from updating while running the code until the end. 

## Summary

### The Advantages and Disadvantages of Refactoring Code
One of the main advantages of refactoring code that we focused on in this analysis was speeding up the code performance. By using different functions to produce the same results, we were able to take up less memory and have VBA read through less code or more efficient code that performs quickly. This can make a drastic difference with extremely large data sets. Another advantage to refactoring is creating a logical and more direct code to read through. For future readers or even if the creator of the code wants to come back to view/adjust the code, depending on how it was refactored, can save a lot of time, and allow it to be easily understood. One advantage I found as a new coder is that I learned different methods of still getting the same results. This will greatly improve my skills and knowledge in this field of study for future purposes. A disadvantage I experienced with refactoring code was the potential for errors or bugs. With my limited knowledge of coding, it was a little difficult to even consider how I could make the original code different, or even better. While refactoring I got stuck and then could not even produce a result until it was resolved. 

### How these Advantages and Disadvantages Relate to our VBA Scripts
The disadvantages gave me problems in my script as I mentioned previously. In more detail, I could not wrap my mind around the other methods of using the limited code that I have already learned. Trying to redo everything I learned but differently caused a lot of errors in my code and I could not get any output. I noticed in my classmate’s work; they had a completely different overall process than I had but their code worked. The issue with their code was that it did not speed up the performance at all. This showed me that if you are refactoring to make a code better through performance time or readability, then it is an advantage. On the positive side of things, while refactoring this code, I dropped some of the “And” conditions that came with the old If-then statements and removed the nested For loop. Truth be told, I needed help from a tutor for this part, but after reviewing the correct code, it made it easier for me to read, break apart, and understand. Of course, as previously mentioned this sped up the code by a lot. 
