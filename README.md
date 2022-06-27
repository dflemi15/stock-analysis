Overview of Project:

The purpose of this analysis is to provide Steve’s parents with a tool that can easily analyze stock data so that they can make the most informed decision on which stocks they should invest in. This tool quickly provides data on critical information such as stock volume and returns to determine the performance of each ticker.

Results:

Stock performance

The portfolio of tickers analyzed overall appears to perform better in 2017 than in 2018. In 2017, the only stock that yielded negative returns was “SPWR” while every other stock included in the analysis provided positive returns between 5% to 200%. In 2018, the portfolio yielded negative returns across the board with the exception of “ENPH” and “RUN.” Although both stocks provided positive returns, the return for ENPH was down approximately 47.6% from 2017 where the return for RUN was up of about 78.5%. 

Run Times

The run times for the original model were approximately 0.85547 when applied to 2017 and 0.85156 when applied to the 2018 data. Conversely, the run times for the refactored model were 0.1289063 and 0.125 for 2017 and 2018, respectively. The refactored model appears to run faster which is expected due to the optimized coding used to construct the macro. Both run times are illustrated below:

![image](https://user-images.githubusercontent.com/107585908/176046634-ddf05eaa-4792-4144-a6e6-4c1a499148bc.png)

![image](https://user-images.githubusercontent.com/107585908/176046646-ea6402d3-0b48-4d8d-9dfc-f53b3288f1d4.png)

Summary:

What are the advantages or disadvantages of refactoring code?

Refactoring code allows you to save time by eliminating redundant work for problems with similar solutions. This gives you the opportunity to optimize the code and tailor it to the specific problem you are trying to solve. There is an opportunity to potentially code something incorrectly which can cause the code to break but overall refactoring can save a lot of time and energy when encountering challenges with similar solutions. In addition, refactoring code allows you to clean the code in a way that makes it easier to read. 

How do these pros and cons apply to refactoring the original VBA script?

The original VBA script had several lines of complex code that included “for loops” and “If, then” statements that would have been time consuming and complicated to create from the ground up for Steve and his parents. By refactoring the code, we were able to replicate much of the complexity of the code and also updating it so that it is now more flexible. Steve and his parents can now seamlessly transition between years of stock data as needed without much difficulty. Also, the refactored code is substantially more easier to read and follow due to the simplicity in which it is currently structured.
