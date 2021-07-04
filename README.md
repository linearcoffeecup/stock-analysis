# **Module 2 Analysis**

## **Overview of Project**

In this project, Excel VBA was used to analyze all stocks in the portfolio.  It was desired to view the volume and the return for each stock.  The database was an excel spreadsheet containing the necessary information:  stock identification, volume, opening price, and closing price.  Two worksheets were created:  “AllStocksAnalysis” and “VBA_Challenge.”  In the “AllStocksAnalysis” worksheet one method was applied, a nested for-loop, and in the “VBA_Challenge” worksheet a method was used with one for-loop and a change of the stock identification within the for-loop.  In each method data needed was gathered.  This data gathering also differed in each analysis in that arrays were used for the starting price and the ending price in the refactored code (“VBA_Challenge).  The analysis time was calculated beginning from when the input was placed into the message box to the end of the VBA program.

## **Results**


|      |**Improvement in Run Time**|             |
| :---: | :--: | :---: |
| Year | Original(seconds) | Refactored(seconds) |
| 2017 | [0.375](https://github.com/linearcoffeecup/stock-analysis/blob/main/Resources/Original_2017.png)| [0.082](https://github.com/linearcoffeecup/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png )          |
| 2018 |.[0.160](https://github.com/linearcoffeecup/stock-analysis/blob/main/Resources/Original_2018.png) | [0.078](https://github.com/linearcoffeecup/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png )          |


## **Summary**

### Advantages and Disadvantage of Refactoring Code in General

A typical definition of refactoring encountered is that it is a practice or process of improving the code and/or database while maintaining the existing functionality.  The advantages of refactoring outweigh the disadvantages of not doing so for certain projects.  External factors may affect the value of refactoring such as whether the code is stable as well as project deadlines.  Examples fo code-proper considerations which affect the value of refactoring are "technical debt" and certain code "smells."

Benefits of refactoring include:

- Optimizing existing code
- Improving code quality
- Future maintenance and enhancements easier
- Lower cost of enhancements
- Code is easier to read
- Speed

Potential drawbacks for refactoring include:

- Complexity
- Long time to complete
- Resource intensive

Advantages and Disadvantages of Refactoring This Analysis

This program had a relatively small data set, and so is not really a candidate for refactoring unless it it running slowly even with a small data set.  The improvement in speed was noticeable for this project, and would be important for larger data sets.   The refactoring took advantage of an improved method, or, rather, reduction of the number of for-loops used.  The only other refactoring feature that is beneficial here is the addition of comments to make the code more readable.  Applying refactoring to this project does not have any disadvantages.
