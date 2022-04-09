# Overview of Project
## The VBA of Wallstreet
### Refactor VBA code and Measure Performance
The second module challenge was to refactor VBA code in stock database to gather stock information in years 2017 & 2018 and determine whether or not the stocks are worth investing. I am refractoring the original code to loop through the data once to collect all information to increase the efficiency.

- [Stock Analysis xlsm](https://github.com/matthallman/Week_2_VBA_Homework/blob/main/VBA_Challenge.xlsm)

### Challenges and Difficulties Encountered
Wrapping my head around for loops and multi-worksheets was a challenge that took much practice.
Encountered numerous runtime errors, but fixed them by going through step by step to find where the hang-ups were.

## Results

![2018 Ticker Results](https://github.com/matthallman/Week_2_VBA_Homework/blob/main/Resources/All_Stocks_2018.png)
We're looking at the years 2017 and 2018 stock returns for certain ticker symbols. 
- 2017 stocks listed all saw positive returns, with the outlier being TERP (-7.2%). 
- 2018 stock returns show negative returns with the exception of ENPH (+81.9%) and RUN (+84.0%).

![2017 Ticker Results](https://github.com/matthallman/Week_2_VBA_Homework/blob/main/Resources/All_Stocks_2017.png)

The refactored code runs much faster then the original code.

## Summary 
### Pros & Cons
Advantages of refactoring the original code are a much cleaner and efficient code as well as a reduced run-time. A user with little to no experience can use the code and get the same results. Using the refactored code allows use of much larger datasets then the original code. 
Using the refactored code could be a small con when using smaller datasets that may not need such complex coding. 
