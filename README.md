# The VBA Green Stocks Analysis 2017 & 2018

## Overview of the project 
The purpose of this Analysis was to Refactor the All Stocks Analysis [VBA Challenge].(https://github.com/justinamaze/stock-analysis/blob/main/VBA_Challenge%20.xlsm). so I can make the code used in that prior analysis more accurate and time effecient. Refactoring helps the with using less memory, taking fewer steps, and improving
improving the logic so the results are easier to read. Then I can determine if the code made the VBA scrip run faster
and easy to utilize for Steve's parents. 



## Results
Code: In the orginal All Stock Analysis I Dim'd the StartingPrice and endingPrice as Single to get an percentage of their value. I had to loop 
through the rows to get the exact prices for the particular tickers. I used the for loop for my tickers and a nested for loop to code 
to get the starting and ending prices placed.

Code: In the Refactored stock Analysis I created a tickerIndex so the number can start at zero. Also instead of using a nested loop as I did in the
Original Analysis I Dim'd the (tickerVolumes, tickerStartingPrices, and tickerEndingPrices.) This particular process allows me to cancel out the nested 
loop and loop threw the Index and Volume

*Original Stocks Analysis Code*

![Resources/Original_Stock_Analysis_Code.PNG](/Resources/Original_Stock_Analysis_Code.PNG)

*Refactored Stock Analysis Code*

![Resources/Refactored_Stock_Analysis_Code.PNG](/Resources/Refactored_Stock_Analysis_Code.PNG)



With the provided picture below I'll start off with the Refactored All Stock 2017 then 2018 Analysis, and give you feed back
from what I see from the photo. In the Refactored All Stocks year 2017 the returns from the 12 companies were mostly positive.
Especially for "DQ" they received a 199.4% return for the year. "Terp" was the only company that was in the negative.
In the All Stocks year 2018 most of the companies took a huge hit except "ENPH" and "RUN". They were the only companies
to have a positive outcome of the year. Although "Terp" was in the negative for both years this was the only company that didn't 
have a dramatic change within the years.

*Refactored year 2017*

![Resources/VBA_Challenge_2017.PNG](/Resources/VBA_Challenge_2017.PNG)

*Refactored year 2018*

![Resources/VBA_Challenge_2018.PNG](/Resources/VBA_Challenge_2018.PNG)

Also in the photos below i have the Run time for the orginal and refactored analysis for both years. 
The time it took to run the Orginal Analysis was (2017 0.609seconds and 2018 0.613secs ). The time it took to
rund the Refactored Analysis was (2017 0.136seconds and 2019 0.167secs). Obviously the refactored analysis
ran faster than the orginal analysis. This tells me editing and refactoring code does work in the user's favor.
I also believe the Refactored code ran faster because it didn't have as much data and coding as the Original analysis
had. 

*Orginial Time 2017*

![Resources/2017_Original_Time.PNG](/Resources/2017_Original_Time.PNG)

*Original Time 2018*

![Resources/2018_Original_Time.PNG ](/Resources/2018_Original_Time.PNG)

*Refactored time 2017*

![Resources/2017_Refactored_Time.PNG](/Resources/2017_Refactored_Time.PNG)

*Refactored time 2018*

![Resources/2018_Refactored_Time.PNG ](/Resources/2018_Refactored_Time.PNG )

## Summary

***1. What are the advantages or disadvantages of refactoring code?***

The advantages of refactoring code is being able to edit and modify the original code to mae it more simple. Also refactoring code
helps save more storage space which leads to a more accessible and faster time to look up the data. The disadvantage of refactoring 
code would be having to figure out what exactly you're going to modify. Also changing code can lead to more errors and debugging issues.
So you'd have to be very careful so you don't mess up the orginal outcome. 

***2. How do these pros and cons apply to refactoring the original VBA script?***

Well the pro to refactoring the original VBA Script the first thing I noticed was the time in the message box was shorter. That had to do with the 
fact that refactoring takes up less data. Also in this instance of refactoring the original VBA script the coding was much simpler. The con to refactoring
the original VBA script would be trying to see going through the hassel of builing another code is worth it. Especially if the original code is working already,
and you have enough storage do I really 'Need' to refactor the script?
