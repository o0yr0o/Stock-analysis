# Stock Analysis with VBA

## Overview of Project

Steve wants to analyze more research for his parents which is to expand the dataset to include the entire stock market over the last few years. He had the original code to loop through the data to collect all the information he needed. However, he would like us to refactor the original code to collect the same information that the original code provided but with more efficient way by taking fewer steps or making the script run faster.

## Results

### Stock Performance
Since Steve wants to analyze the entire stock market in the past few years for his parents, we can use this graph to visualize and make him and his parents have a better understanding of the stock market in 2017 and 2018 .

![Stock_Performance_Comparison](https://user-images.githubusercontent.com/82549782/117394195-b4cb2500-aec3-11eb-909a-4cddec1c42f2.png)

The total daily volume is the total number of shares traded throughout the day, so the total daily volume for one year can provide a rough idea of how actively stocks was traded in 2017 and 2018. The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year. In other words, it is how much your investment grew by the end of the year.

According to the graph, although DQ's yearly return in 2017 was very high, it fell to the lowest return in the entire stock market in 2018. DQ's total daily volume is also low, despite the 2018 trading volume has increased, but still far less than other stocks. So I would not recommended this stock to Steve's parents.

The yearly return of the entire stock market has declined in 2018, but only RUN and ENPH have increased. The annual return of RUN increased from 5.5% to 84%, and the daily trading volume in 2018 also increased, indicating that the stock has a continuing upward trend. ENPH, on the other hand, had a positive return in 2018, but its annual return fell from 129.5% in 2017 to 81.9% in 2018. And his daily transaction volume has grown rapidly from 2017 to 2018. This suggests that the stock also has a tendency to rise, but once future daily trading volume drops, its annual return is likely to fall even faster. Therefore, comparing the two stocks, I would recommend Steve's parents to invest in Run, but if his parents can tolerate higher risks, ENPH can also consider investing.

### Comparison of Original Code and Refactored Code
The main purpose of code refactoring is to refactor the code into a more simplified or more efficient form. Refactoring code improves its internal structure, thereby increasing execution efficiency, while leaving the external functionality unchanged(Watts & Kidd, 2018). 

#### Execution Times Comparison
![Excecution_Times_Comparison](https://user-images.githubusercontent.com/82549782/117507008-7f224c80-af54-11eb-98cc-e6818f3f78cb.png)

By comparing the execution times between the original script and the refactored script, we can see that the speed to run the refactored script is much faster than the original script. The longer execution times may caused by the nested loop in the original script.

#### Difference between the Original Script and Refactored Script
![Script_Comparison](https://user-images.githubusercontent.com/82549782/117512617-f3adb900-af5d-11eb-84bb-fa6d2d127779.png)

The nested loop was used in the original script. The inside loop first looped through each row of the database to find the same ticker, and calculates the total volume and starting and ending prices of each ticker. Then, the 12 tickers are looped in the outiside loop to output the data of each ticker. 

In order to reduce the execution time, the solution to avoid the nested loop is to index something.   Therefore, ticker index is used to replace the inside loop. When I created the three ouput arrays, I put a bracket of 12 to tell the computer that the highest of tickerIndex is 12. So, the second loop is to loop over all the rows in the spreadsheet, but because tickerIndex is involved, the computer would loop 12 times to calculate the total volume and starting and ending prices for each ticker. 

## Summary

### Advantages and Disadvantages of Refactoring Code
- Advantages (Cuelogic, 2021):
  - Refactoring improves the design of software.
  - Refactoring makes software easier to understand.
  - Refactoring helps finding bugs and programming faster.

- Disadvantages:
  - Refactoring required more money and more time.

### Advantages and Disadvantages Apply to Refactoring the Original VBA Script
- Advantages:
  - The refactored code runs faster.
  - Helps me understand more clearly the purpose of the original code and how it works.

- Disadvantages:
  - Takes a long time to refactor.
  - New bugs might appeared.

## References
Cuelogic. “What Is Refactoring and Why Is It Important?” _Cuelogic Technologies Pvt. Ltd._, Cuelogic Technologies, 17 Jan. 2021, www.cuelogic.com/blog/what-is-refactoring-and-why-is-it-important. 

Watts, S., & Kidd, C. “What Is Code Refactoring? How Refactoring Resolves Technical Debt.” _BMC Blogs_, 19 Mar. 2018, www.bmc.com/blogs/code-refactoring-explained/.
