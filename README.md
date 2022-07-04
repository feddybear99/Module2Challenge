**Module 2 Challenge** 





**Overview of Project**
  
The purpose of this project was to help Steve help his parents find the total daily volume and yearly return for multiple stock tickers. Their focus was to try and compare / contrast multiple tickers value by examining each tickers yearly returns as a measurement of it’s potential investment value. The analysis will be done in Microsoft Excel by enabling macros and creating a program flow that loops through all of the 2017 and 2018 (stock) tickers to find which stock(s) have a positive rate of return for investment valuation. 

**Results**

In the screenshot below, <img width="1171" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/106992995/177092195-3112bf18-ef2b-4359-8e4c-8c8d9722e8ee.png"> the excel spreadsheet was able to organize and sort specific rows / columns through VBA. The popup in the middle of the screenshot is meant to explain how fast the results were able to pull up for each years’ stock data. In this case, the code ran in 0.078125 seconds for the year 2017. It’s important to note that finding the time was only possible by using the “Timer” capability within VBA. Steve’s parents were concerned with pulling up large amounts of data, and wanted the time stamp in each to measure efficient run times. 
Due to conditional formatting written within the VBA script, Steve’s parents are also able to see at a quick glance which stock tickers were successful (green highlight) or not successful (red highlight) in either 2017 or 2018 ticker return. In the 2017 ticker run, it’s output shows that all tickers returned successfully EXCEPT the ticker ‘TERP’, which had -7.2% return in 2017. 

In the next screenshot, <img width="1069" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/106992995/177092256-cd90228a-6b6e-46e5-a97e-1945e7581db2.png"> the excel spreadsheet popup in the middle of the screen ran in 0.078125 seconds for the year 2018. These results returned the same amount of running time (in seconds) it took the computer to find and organize the data on the output sheet. Since Steve’s parents mentioned they were primarily concered about finding which tickers generated high yearly returns, it’s easy to see that in this data set, a lot of tickers at first glance are red - meaning “negative return”, which is what Steve's parents want to stay away from. 

Honestly, if Steve’s parents want to invest in stock, they should choose the tickers that resulted in positive stable growth for *both* 2017 and 2018. This would mean that both ENPH & RUN tickers would be the priority tickers to focus on from an investment perspective. 


**Summary**

*Advantages of refactoring code:*
The acronym DRY: Don't Repeat Yourself is important to follow because it reduces reusing a piece of code over and over again. I noticed when there were multiple places where I needed to make the same edit, and that’s where it led to redundancy and where I needed to correct each fix —which resulted in several troubleshooting problems.

*Disadvantages of refactoring code:*
It can be very hard to work off of a previously built module (in Excel) because of the way VBA is set up. I had to re-do quite a few Modules within VBA so that the code would eventually work just right. I found that it was hard to rearrange and fix dimmensions for the code when trying to run a for loop and an array. 

*Advantages of original VBA script:*
It was straightforward when I first built it for the “DQ analysis” because it only had 1 loop for the output and did not have to go through multiple tickers to find multiple values. With this new VBA script, the major difference was creating an array for a handful of tickers and organize the output data in a visual pleasing way. 

*Disadvantages of refactored VBA script:*
Making sure the right worksheet is active in VBA. It was hard to create loops and nested loops while also remembering to “Activate” a particular worksheet so that the correct loop could finish its cycle. 
