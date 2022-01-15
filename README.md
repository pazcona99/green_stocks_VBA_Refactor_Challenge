# VBA of Wall Street: Refactoring VBA Code & Measuring Stock Performance
Using VBA  to refactor the stock data that can loop through the itself one time and collect all of the return information. 

## Overview of Project

In this challenge, a code was provided to analyze a subset of data associated with stocks, their respective prices, and returns. While only at a click of a button this information can be provided to the analyzer, the purpose of refactoring the existing code is to make it read through thousands of stocks, not just specific subsets. This is important to help create efficieny in analyzing information and code run time. For example, the original code was set up to anaylze a dozen stocks. While the length of this array may change, depending on the need of the anaylsis, it would be a time consuming process to continue to modify the code to the specific stocks required. Rather, it is more efficient on the users end to refactor one code such that it is not dependent on initial stock ticker inputs, and can still provide the same information regardless of the size of the data. This would also allow for the code itself to run more smoothly, and therefore create a smaller run-time and taking up less memory.

Refactoring in of itself is a crucial aspect to working with code. The process strives for efficiency by making it take fewer steps, using less memory, and improving logic. The code modifications would also allow for future readers to easily follow the steps behind the logic. The code is not a static entity as it can always be changed for the better.

---

## Analysis and Challenges

### Analysis of Original Script

For this analysis, two worksheets were provided with data of the associated with certain green stocks. Each ticker had an associated date for which it then include other values such as the opening price, highest price, lowest price, close price, adjusted close price, and volume. For this particular analysis, the goal was to develop a table showing a stocks total daily volume and net return of their value as a percentage.

The code was strutured such that any date within the worksheet can be inputted and then run. After the year is inputted, a timer begins so as to track the time to complete the analysis. An array was initialized with 12 ticker inputs of various green stocks that the analyzer had particular interest to. The code then proceeds to find for every value of that ticker name the respective volume on that particular day. It continues this until the name of the ticker is not found in a succeeding row. Similarly to daily volume, the code will also store the starting and ending price of the particlar ticker it's currently looking for.

The output of this code will be the stock in the form of it's ticker identification followed by the total volume and the return on investment based on the quotient of the ending price and starting price. Once the code has finished it will provide a message box indicated how long the particular instance of code took to run.

Two different datasets of different years, 2017 and 2018, were assessed for their overall volume and return. The figures below illustrate the output of the stocks in the form of a table. 

<img src="Resources\data-M2-Challenge-02-2017-stock-analysis.png" width=500 align=center>


Stock Analysis (2017)

<img src="Resources\data-M2-Challenge-02-2018-stock-analysis.png" width=500 align=center>


Stock Analysis (2017)

The challenge associated with this particular activity is derived mostly in the ability create the correct logic within the `for` loop. This comes with an understanding of what the table is supposed to look like and how the ticker reads each row in the Excel file. One should understand the process of going through each of the rows, and verifying if the current row is the first or last row of it's own ticker in order to move on with the analysis for all stocks. This part of the code is a bit crude because an array was hard-entered with initial inputs specific to the interest of the analyzer. 

Once this code was run, for each year, the code outputted how long it took for each session to run. The figures below indicate that the year of 2017 and 2018 took approximately **0.8203 seconds and 0.8438 seconds**, respectively. These time values are very close to each other within two sigificant figures, but the float is long on the actual value shown.

![Original VBA (2017)](Resources\Original\VBA_Module_2017.PNG)

Original VBA Time Stamp (2017)


![Original VBA (2018)](Resources\Original\VBA_Module_2018.PNG)

Original VBA Time Stamp (2018)



### Analysis of Refactored Script

For this analysis, the code was refactored to make it easier for it to accomplish the same task but not constrain it by the number of stocks. The purpose was to make the code flexible enough so that it can read through more than just a dozen stocks and produce the same output. To do this, the code was modified to have a `tickerIndex` varaible that was initalized to zero before iterating through all the rows. The purpose of this index is for the code to access the correct index across different arrays.

Through the use of a 'for' loop and 'if then' statements similar to the original code, this code was able to provide the same output arrays for the ticker volumes and return on investment. In running the stock analysis, it was confirmed that the outputs for 2017 and 2018 were the same as the original code. However, the time spent running the code was slightly faster as shown in the figures below.

![Refactored VBA (2017)](Resources\Refactored\VBA_Challenge_2017.PNG)

Original VBA Time Stamp (2017)


![Refactored VBA (2018)](Resources\Refactored\VBA_Challenge_2018.PNG)

Original VBA Time Stamp (2018)

Once this code was run, for each year, the code outputted how long it took for each session to run. The figures below indicate that the year of 2017 and 2018 took approximately **0.8203 seconds and 0.8438 seconds**, respectively. These time values are very close to each other within two sigificant figures, but the float is long on the actual value shown.

---

## Summary

### Advantages and Disadvatanges of Refactoring in General
Based on the afforementioned code anaylsis, it would seem that refactoring code provided a case for increasing the efficiency in it's runtime. There are many advantages to refactoring code including but not limited to making it more neat and clear for future maintenance, improved bug fixes and performance, enhanced security and structure, and scalability. This optimization of the code allows for the same end result without adding new functionality. The downstream benefits of refactoring are almost immediate with the advent of having saved time and costs as there is less likelihood of errors to occur or continuous modificaton of the code itself. 

There are some, though few, disadvantages to the practice of refactoring in general. Refactoring could be a risky procedure especially when the application is big. Proceeding with refactoring in such a big application may make it more prone to errors that may or may not be accounted for. Refactoring only works when the proper test cases are available, but when this doesn't exist for the code to truly work then it can be less than the originally functional. There is also inherent risks in refactoring code espeially when the developers do not understand the purpose of the particular application. This could lead to more errors than what was previously mentioned.

### Advantages and Disadvatanges of Original and Refactored VBA Script
One advantage of the original script was that it was pretty straightforward to understand the array being created because it was tailored for specific tickers of interest. The array was initialize for the dozen tickers that the code was to seach for which made the indexing simple. However, this also served as it's own disadvatnage since the tickers initialized were limited to only a dozen. While more or less tickers can be added to the array, the process of doing that each time the code is going to be run is a time consuming process and nulls the entire point of creating a algorithm do this tedious work for the user. Another disadvantage to this code is that it takes a but longer to run because it is searching at least 200 times for each ticker match. This computing process time may seem insignificant, but in the grand scheme of being provided thousands of data points, it will add up.

The advanatges of the refactored script are quite clear in that it allows for greater efficiency without compromising the functionality of the code. The logic can be made more clear so that futre developers can test it and understand it at the same time. This impacts the computing process time that in effect will allow for less RAM to be used, less energy spent on computing, and so less money spent overall. The refactored script uses an index initialized at zero that builds on itself as it loops thorugh the conditionals, which makes it a lot easier for the analyzer to see data regardless of which tickers they actually want to or not already listed. While there are inherent risks in starting the process of refactoing code, a main disadvantage of the refactored script is that the logic was more sound yet difficult to put into practice. Without a proper understanding of the intent, a refactored code may lead to more bugs than actual enhancements which is counter to everything exoected of a refactored code.


