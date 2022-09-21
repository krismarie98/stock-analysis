# Module 2 Stock Analysis Refactoring

## Overview of Project:
### The purpose of this project was to take the data given and summarize. The original data that was given included two worksheets both representing different years. In those worksheets were numerous items of data representing the different tickers. Throughout the module the goal was to use tools like for loop and conditionals to create a new sheet summarizing the data. For example in our sheet DQ Analysis I created a script that showed the total daily volume for the year 2018 of the ticker "DQ". Additionally, in the same module we created a macro AllStocksAnalysis which was alble to list all the tickers,their total daily volume and their returns. This allowed me to practice using nested loops. This all led to the challenge 2. In challenge 2 the goal was to refactor the AllStocksAnalysis so that it was more efficient.
## Results:
### By looking at the time it takes to run the refactored code and the time it takes to run the orignal we can tell the orignal code takes longer than the new code. For example, in the case of 2018 refactored the time was 0.234375 seconds, the original run time for 2018 was 0.5507813 seconds. Additionally you can see that in the year 2017 there were more stocks in the positve than in the year 2018. In 2018 the only stocks not in the negatives for returns were ENPH and RUN. In 2017 every stock had positve returns except for the stock TERP.
![image](https://user-images.githubusercontent.com/112527054/191594918-2138bd8f-a256-4052-9038-da15303b6114.png)


![VBA_Challenge_2018](https://user-images.githubusercontent.com/112527054/191595284-6421bac0-0ce5-4b52-9196-8d615f2f5d45.png)

![Refactored table 2017](https://user-images.githubusercontent.com/112527054/191596840-716c6ac3-56e1-44dd-92de-806667fc8584.png)

![Refactored table 2018](https://user-images.githubusercontent.com/112527054/191596899-970158f5-1599-4ffc-872c-446187340158.png)

## Summary:
### 1.What are advantages or disadvantages of refactoring code?
#### The disadvantage of refactoring code is that is can be harder to understand. Though the goal is to use less steps the techniques that are used can be harder to understand. An advantage is that the code may run faster.
###   2. How do these pros and cons apply to refactoring the orignal VBA script?
####     A pro of refactoring the orignal VBA script is that the script now runs faster than the orignal. A con is that it was easier to read the orginal code than the new one.
