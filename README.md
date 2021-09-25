#Green Stocks – VBA Challenge


##Overview of Project:

Steve, our client, is looking to analyze data for this parent’s investment. The purpose of this project is to analyze stock data to show the different stock trends year over year and the returns each stock yields. Original VBA macro ran data for twelve stocks, refactoring the VBA macro is necessary in order for the VBA code to run smoothly and more time efficiently if there will be hundreds of stocks to analyze. 

##Results

![All Stocks 2017 ](https://user-images.githubusercontent.com/89940569/134753177-143916fb-0c99-499c-861c-775a801e253e.png)
![All Stocks 2018 ](https://user-images.githubusercontent.com/89940569/134753180-631cc938-483c-4940-8541-8376495f433b.png)


When evaluating stock performance years 2017 to 2018, one can identify that eleven of the twelve stocks in the year 2017 yield a positive return, however in 2018 there are only two stocks which yielded a positive return. Based on the year 2017, there are many profitable investment choices, stocks highlighted in the green represent the positive increase from the beginning of the year to last day of the year. In 2018 we see a significant downshift in stock return, highlighted in red, ten of the twelve stocks were negative yielding. Based on the data gathered one can conclude that all stocks decreased in value, with exception of “RUN”. “RUN” is the only stock analyzed which demonstrated year over year growth. 


The original VBA Macro running time was longer than the refactored data total time. Although refactoring data did not make noticeable time difference for our current set of data being analyzed as they both were under 1 second, it would be a noticeable time difference if we add more lines to this data set. If an exponential amount of data was to be added to out dataset to be analyzed, on average the 2017 data would run 77.5% faster using the refactored macro vs the original, and the average 2018 data would run 61.6% faster using the refactored data. 

Original VBA Macro
2017 - .6953125   2018 - .703125

Refactored VBA Macro
2017 - .15625   2018 - .2695312


<img width="328" alt="GreenStocks 2017 " src="https://user-images.githubusercontent.com/89940569/134753206-6d9b2fa1-6202-4412-8b15-7beb44078854.png">

<img width="328" alt="GreenStocks 2018" src="https://user-images.githubusercontent.com/89940569/134753209-a9de6ec1-dbe4-43b3-83b8-7a7b32331c50.png">

<img width="372" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/89940569/134753221-1dabda5b-892d-4766-8faf-e7990869ffd0.png">

<img width="372" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/89940569/134753231-2c0b9eef-d9c6-42fe-bf97-832a19ba65f4.png">


##Summary

There are many advantage to refactoring VBA Macro data such as decreasing run time for a macro with a large dataset and revising the macro to perform more efficiently. Refactoring allows for macros to be run more effectively and efficiently. When analyzing large data sets, refactored macros are able to “multi-task” in order to produce the same results in a shorter amount of time, it helps eliminate overlapping of loops over the same lines of data. Disadvantage of refactoring data is that it may lead to extensive debugging. Debugging may be necessary in order to fix any macro composition which may no longer run smoothly once modification to the original macros are made.  It is wise to revise written VBA macro to be as efficient as possible, however the macro must keep its integrity and ability to achieve the desired analytical result. 



