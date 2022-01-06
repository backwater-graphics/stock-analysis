# VBA Stock Analysis with Excel
## Overview of Project
My friend Steve just graduated with a finance degree, and his parents wanted to be his first client. They know that they want to invest in green energy and believe that the future will have a reliance on alternative energy products. They have not done a lot of research on the topic but they know that they want to invest in Daqo with a ticker name DQ since Dairy Queen or DQ was where they meet. Steve was concerned by this because he believes that they need to have a diversified portfolio. Steve wants to do the Research but knows that there is a vast amount of information and will need help with this task. Steve asked if I could use Excel to write a VBA script to analyze stock data to find the best option for Steve’s parents.
### Purpose
The purpose of this challenge was to create an efficient way to look at multiple stock tickers from 2017 and 2018 to determine if the stocks are worth investing in for Steve’s parents by using a VBA macro. The project needed to have a macro created to analyze the stock tickers and then it called to have that macro refactored to see if we could get the macro more efficient, which would allow our client Steve a better way to analyze each stocks performance over a one-year period.
## Results 
The analysis displayed on the tables below are for about a dozen green-energy stocks which the companies are known for using one of the following alternative energies in the form of either hydroelectricity, geothermal energy, wind energy, or Bio Energy. The table contains three groups of data:
   •	Ticker name
   •	Total daily volume for a given year
   •	Percentage of a yearly return
###Daily Volume and Returns Charts
###Comparison of stock performance between 2017 and 2018 
As we review the Daily Volumes and the Return charts for both 2017 and 2018, we can see that the green stocks we analyzed performed much better in 2017 than they did in 2018 in terms of the yearly returns and daily volumes with the exception of ENPH and RUN. In both years ENPH and RUN both had positive returns and also increased their daily trading volume.
Since Steve was looking into the stocks for his parents, we also looked at the results of the DQ stocks in 2017 which had a low volume and a high yearly return (at that time this might have been an indicator of a company on the rise). However, the situation of DQ stocks in 2018 had changed completely. That stock closed its year with negative 63%. The trading volume was higher, yet it didn’t result in a positive outcome. The results of this analysis confirmed that DQ stocks would be risky investment for Steve’s Parents.
##Code Comparison for Execution Times
In the original code (see below) I used nested loops which run through each record of the data set once and with a small number like our challenge of 12 records it is fine, but if you had a data set with several thousand records it would strain the system by running through the values n2 times.
 
 I knew that if I wanted to reduce the time and the resources consumed, I would need to refactor the code into an array (see below), by doing this the computer now only needs to access each data row once, because of the data sets are presorted into essentially a zone and once its left that zone it does not have to check back through those rows to confirm that there are no other values in that range. This is what helps to cut down the time resources required to n, rather then the original code which requires n2   for the calculations. 
 
By refactoring the script, we can see that it improved the execution times from the original script. See examples below for the comparison of execution times.
###Original Code Execution Times
  
###Refactored Code Execution Times
  
## Summary
In a summary we have discovered what the advantages and disadvantages of refactoring the code has on VBA scripts. Below I have listed these out:
###Advantages of refactoring code:
•	By refactoring the code it helped to make it readable or easier to understand and make the code clearer by the improved logic of the code.
•	It also helps to fix any bugs that may have been overlooked in the original code, which helps to improve the functionality. and reduce the complexity
•	Refactoring the code helps to reduce the complexity by restructuring the code to take fewer steps which lowers the amount of memory the computer uses to do the calculations and therefore takes less time to execute the code
###Disadvantages of refactoring code:
•	One of the biggest disadvantages to refactoring code is that it can be time-consuming
•	Refactoring may cause the outcomes to be altered or it may even cause and already working code to be destroyed by errors in the refactoring code
•	Another issue is that by refactoring the code, we could end up with a less efficient script
•	Refactoring can be frustrating at times as the code may not be well commented so you spend a lot of time figuring out what each line or block of code supposed to do.
•	Refactoring can also be tuff to do based on the size and complexity of the code. 
###How does these pros and cons apply to refactoring the original VBA Script?
For this challenge refactoring the code doesn’t require adding any new functionality, except for trying to improve the timing of the code speed. First let’s get started with how the pros apply to refactoring the original VBA Script, when we revisited the code, we had new ideas which helped to improve the overall process of the code and helped to cut the run time by a good percentage of the original time.  It also helped to cut some of the complexity of the original code by replacing the nested loops with the arrays.
For how the cons apply to refactoring the code is that it was frustrating to refactor the code. The first time I did this it gave me the following error “Error 6 Overflow” I tried a few things still nothing changed so I brought it to the class office hours which I got the same error and also another error “Error 9: subscript out of range”. This was very time consuming, but I went through the whole code line by line and figured out I had a piece of code out of order when I rearrange it, it worked. This was the only con that I found.
