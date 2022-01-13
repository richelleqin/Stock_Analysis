# Stock Analysis

## Overview of the Project

### Purpose
The purpose of this analysis is to refactor the code so that it is efficient and optimized to run for a large number of stocks. Further, we want to provide analysis to compare the stock performance between 2017 and 2018.

## Results

### Analysis
In general, there are positive returns in 2017 for the stocks except for TERP. In 2018, there is an overall loss except for ENPH and RUN. DQ had to greatest % change of -262.0%. 
Before refactoring, the code ran for 0.890625 seconds for 2017, and after refactoring it ran for 0.34375 seconds for 2017. 
![2017](https://user-images.githubusercontent.com/67567087/149422128-b041da2e-1820-4eb0-a959-1a33c739e5d1.PNG)

Before refactoring, the code ran for 0.8828125 seconds for 2018, and after refactoring it ran for 0.328125 seconds for 2018.
![2018](https://user-images.githubusercontent.com/67567087/149422256-c57bb4a5-3ab9-4cbe-8793-8500f6e707e0.PNG)

## Summary 

### Advantage and Disadvantage of refactoring code in general
The advantages to refactoring code includes an increase in accuracy, efficiency, and usability of the code. This means that the refactored code can be better maintained and better read. The code should decrease in complexity and should be refactored to be scalable should something change. The code itself should also run faster as it is more optimized for the job. 
The disadvantages of refactoring code are that during the refactoring portion, information may be missing. Validation is extremely required and hard to do as you are rewriting portion of code. Further, without knowing the scope for the future, it is hard to test the code’s scalability for that future. 

### Advantage and Disadvantage of the refactored VBA script
In the original code, ending price and starting price assumes that the stock will always be in order and sorted. This is not scalable, and the code will cause inaccuracies. Further the usage of arrays to store values allows better future validation for maintenance. If something should go wrong, we do not need to rerun the code to the specific stock to get the value, instead we can just call the tickerIndex associated to the stock. The code is also more efficient.
The cons for this is that during refactoring, it is hard to validate the code’s correctness until after the entirety has been reprogrammed. As the variables and the iteration of the code changed, as well as the logic behind tracking and printing the values, until the code is completely refactored, the values could not have been compared. If there were to be issues, it is hard to tell exactly at which portion of the refactoring that issue may have occurred, as we could not have compared step by step while programming. 
