# An Analysis of Stock Tickers
___
# Purpose:
Performed data analysis on a set of tickers to identify which tickers had the best and worse performance in 2017 and 2018.   Also performed an analysis on the run time of the original VBA script and refactored VBA script.

# Analysis:

### **Analysis of Tickers in 2017**
![2017 Tickers](https://github.com/frlinh/stock-analysis/blob/c62fd35b6916e481665cf89bc5df2693b5085e57/VBA_Challenge_Tickers_2017.png)

### Performance
Original VBA script: Analysis ran in 0.81 seconds

Refactored VBA script: Analysis ran in 0.08 seconds

![2017 Run Time](https://github.com/frlinh/stock-analysis/blob/c45e74f3a7a53e77f367eb688aadfdf18b6c9a29/VBA_Challenge_2017.png)

### **Analysis of Tickers in 2018**
![2018 Tickers](https://github.com/frlinh/stock-analysis/blob/8d4087f4387c745b9e9f41788f41606823d41f49/VBA_Challenge_Tickers_2018.png)

### Performance
Original VBA script: Analysis ran in 0.79 seconds

Refactored VBA script: Analysis ran in 0.06 seconds

![2018 Run Time](https://github.com/frlinh/stock-analysis/blob/8d4087f4387c745b9e9f41788f41606823d41f49/VBA_Challenge_2018.png)


### **Summary**
The refactored VBA script is less complex, easier to read, and simpler to manage.  Reducing the run time will be extremely helpful when we run analysis on larger datasets.  The disadvantage when setting up refactored VBA script is when the code is not setup correctly, the outputs would be incorrect, which can cause bugs and errors in the script.


The refactored VBA script ran more efficiently than the original VBA script because the refactored VBA script loops over the array for outputs in comparison to the original VBA script which loops over rows for outputs.  This is essentially more effective because the refactored code simultaneously outputs the Ticker, Total Daily Volume, and Return, while the original VBA script outputs the Ticker, Total Daily Volume, and Return for each row, then runs the output again for the next row.
