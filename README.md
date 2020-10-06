# Analyzing Stocks with VBA

## Overview of Project

### Purpose

The purpose of this analysis was to analyze stock data using VBA in order to see which stocks would be a good investment. Through this exercise we were able to learn how VBA can be utilized to analyze large quantities of data very quickly. We were also able to see how code can be written in different ways, also known as code refactoring. While different versions of code can produce the same result, the efficiency and resources used to do so can vary greatly. 

## Results

Thanks to the formatting being performed in our macro, we can cleary see that 2017 had far better stock outcomes than 2018. As you can see below, in 2017 of the 12 stocks only one had a negative return. On the other hand, in 2018 all but two had negative returns. Now to compare the original macro code with the refactored code, the original code performed much better. As shown below, the time it took to run the refactored code took about 7-8 seconds to run as opposed to 1 - 1.5 seconds to run the original code. In the original code, we used nested loops to get our result. In the refactored code, we used arrays and non-nested loops to get the result. Perhaps storing the values in arrays rather than just accruing the value in a variable resulted in such a performance difference.

## Summary

- What are the advantages or disadvantages of refactoring code?

One advantage of refactoring code is that you can make code more efficient which could improve the performance. Another advantage is you can make code more reusable by replacing hard coded values with variables. An advantage could also be for the person doing the refactoring; the act of rewriting code to do the same thing can be a valuable learning experience.

A disadvantage of refactoring code is that you can reduce the efficiency of code and worsen performance, which is shown with this exercise. A less experienced or knowledgeable developer coming in and worsening code is a risk. Another disadvance could arise if the original code didn't have clear comments and the person refactoring the code changes what the code is actually doing.

- How do these pros and cons apply to refactoring the original VBA script?

The pros of refactoring the original VBA script were that we could see how we can achieve the same outcome different ways. The con to that was that in this instance refactoring actually made the code less efficient. However, it was a good learning experience to see for ourselves how the decisions we make when writing code can effect performance. This encourages new developers to be thoughtful and avoid lazy coding.
