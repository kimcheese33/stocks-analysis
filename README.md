# Analyzing Stocks with VBA

## Overview of Project

### Purpose

The purpose of this analysis was to analyze stock data using VBA in order to see which stocks would be a good investment. Through this exercise we were able to learn how VBA can be utilized to analyze large quantities of data very quickly. We were also able to see how code can be written in different ways, also known as code refactoring. While different versions of code can produce the same result, the efficiency and resources used to do so can vary greatly. 

## Results

  Thanks to the formatting being performed in our macro, we can clearly see that 2017 had far better stock outcomes than 2018. As you can see below, in 2017 (see Exhibit A or B) of the 12 stocks only one had a negative return, shown clearly in red. On the other hand, in 2018 (see Exhibit C or D) only two had positive returns, shown clearly in green. Now to compare the original macro code run time with the refactored code run time, the original code performed slightly better. As shown below, the time it took to run the refactored code took between 0.700 to 0.711 seconds (see Exhibit B or D) to run as opposed to 0.660 to 0.684 seconds to run the original code (see Exhibit A or C). 

  In the original code we had one array to store the ticker names and a nested for loop. The inner loop of the nested loop looped through the data in the worksheet to accrue the total volume value in a variable as well as gather starting and ending price in their respective variables. The outer loop loops through the ticker array and outputs the values from the inner loop to the appropriate cell in our result sheet. In the refactored code, we used four arrays to store our values: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. To tie these four arrays we used a common index, tickerIndex. To get our values, we also utilized a nested for loop with a similar format to the one we used in our original code. The inner loop looped through the stock data accruing and storing values in our arrays instead of variables. The outer loop was used to output the data onto our "All Stocks Analysis" sheet. 
  
  The main difference between our original code and our refactored was the manner in which we were storing our values. In our original code we were accruing them in variables while in our refactored code we were storing them in arrays. Our run time result showed that the original code ran slightly better. The reason for that could be that storing values in an array vs. a variable could be less efficient in this instance. Although the run time difference seems insignificant, if we were to run this script over and over again and/or for much larger datasets the difference could add up to be significant. As a result, might be best to stick with the original code. 

- Exhibit A: 2017 Results Using Original Code
![image_name](https://github.com/kimcheese33/stocks-analysis/blob/main/VBA_Challenge_2017_Original.png)

- Exhibit B: 2017 Results Using Refactored Code
![image_name](https://github.com/kimcheese33/stocks-analysis/blob/main/VBA_Challenge_2017.png)

- Exhibit C: 2018 Results Using Original Code
![image_name](https://github.com/kimcheese33/stocks-analysis/blob/main/VBA_Challenge_2018_Original.png)

- Exhibit D: 2018 Results Using Refactored Code
![image_name](https://github.com/kimcheese33/stocks-analysis/blob/main/VBA_Challenge_2018.png)


## Summary

- What are the advantages or disadvantages of refactoring code?

One advantage of refactoring code is that you can make code more efficient which could improve the performance. In real world situations, time/resources = money. Thus, making sure code produces the desired result utilizing the least amount of time and/or resources is ideal. Another advantage is you can make code more reusable by replacing hard coded values with variables. For example, if we see a piece of code being used over and over again, refactoring the code to make it reusable will reduce complexity and overhead. An advantage could also be for the person doing the refactoring; the act of rewriting code to do the same thing can be a valuable learning experience. 

A disadvantage of refactoring code is that you can reduce the efficiency of code and worsen performance, which is shown with this exercise. A less experienced or knowledgeable developer coming in and worsening code is a risk. Another disadvantage could arise if the original code didn't have clear comments and the person refactoring the code changes what the code is meant to do. Code can be complex and difficult to understand. If the original developer didn't add comments to indicate the purpose of his/her code, someone else could come in later and misinterpret what he/she was trying to do.

- How do these pros and cons apply to refactoring the original VBA script?

The pros of refactoring the original VBA script were that we could see how we can achieve the same outcome different ways. The con to that was that in this instance refactoring actually made the code a little less efficient. However, it was a good learning experience to see for ourselves how the decisions we make when writing code can effect performance. This encourages new developers to be thoughtful and avoid lazy coding.
