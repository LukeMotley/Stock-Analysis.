# Stock-Analysis.

## Overview of Project:

The purpose of this project is to refactor a prexisting VBA code to shorten the computation time of the macro. This was done by rewriting code to be more efficient and clear in its descriptions while also remaining scaleable to the needs of the user. The macro takes several metrics of a list of stocks and displays the total volume of that stock and the return on investment. In the current data set, there are 12 stocks.

## Results:
  The 2017 and 2018 original run times are 0.6367188 seconds and 0.6484375 seconds, respectively. This is while computing a dataset of 3013 rows, 8 columns, coming out to 24,104 cells. The Refactored code, with the same data took 0.1054688 seconds and 0.09765625 seconds for 2017 and 2018, respectively.!
   
 **The refactored code is about 6 times faster than the original.**

### Original Macros ###

![Original macro 2017](https://user-images.githubusercontent.com/85656361/126077251-3bfdf5ba-d6ce-4fd3-ba8e-db432a859a44.png)

![Original macro 2018](https://user-images.githubusercontent.com/85656361/126077252-b38925d0-145c-4cdc-b98b-5b36a07e28f8.png)



### Refactored Macros ###

![VBA_Challenge_2018 png](https://user-images.githubusercontent.com/85656361/126077228-649c8f1e-99ec-4d20-acb9-92db7b461c8d.png)

![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/85656361/126077225-4315f88a-1ce4-41ad-8bba-d783077dd412.png)

## Summary: In a summary statement, address the following questions.
### What are the advantages or disadvantages of refactoring code?
The advantages of refactoring code are the potential exponential time that can be saved when running code. Because this particular dataset was relatively small, it doesn't really matter if the macro takes 1/2 a second or 1/10 of a second. But, when running large amounts of data, the time can rack up quite quickly. If the data set was 100 or 1000 times bigger the difference in less than half a second could expand out to being the difference between 10 seconds and a minute of wait time just for the macro to run.
  There is an added benefit to refactoring code. It gives a second opportunity to scan through it and make it more clear as to how it is working. Commenting and formating everything to be cleaner paves the way for other people that may need to use or work with the code.
  That being said, if dealing with small data sets, it may not be worth the time to refactor code. The time that it takes to refactor code can easily outway the small amounts of extra wait time especially if it is a small data set. Refactoring code can also be difficult to see where opportunities to save computation time. 
### How do these pros and cons apply to refactoring the original VBA script?
  If this dataset were bigger, it would be beneficial to refactor the code to increase efficiency. But, because of the requirements needs to analyze these stocks, it is not neccessary to refactor. The biggest benefit of refactoring this code would be to make sure the code is clear and readable so that other people can use it without being confusing. 
