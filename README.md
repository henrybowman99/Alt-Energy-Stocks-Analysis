# Alternative Energy Stock Analysis

## Overview of Project

Utilizing daily data on 12 stocks in the alternative energy industry from the years 2017 and 2018, I sought to create a table displaying each stocks total volume traded for the respective year as well as each stocks annual return. With this information, I can draw conclusions on the past performance of each individual stock while also gathering a synopsis of the overall performance of the entire alternative energy sector.

## Results

As you can see from the All Stocks (2018) and All Stocks (2017) workbooks below, the group of stocks collectively faired far better in 2017 compared to 2018. This conclusion is most evident by seeing that 11 of the 12 stocks experienced positive returns in 2017 while just 2 of the 12 stocks increased in value during the year 2018.

![All Stocks (2017)](https://user-images.githubusercontent.com/95651156/149551820-91785521-bc4d-446e-aa74-4f68a7c2a90c.png)
![All Stocks (2018)](https://user-images.githubusercontent.com/95651156/149551846-efc8e5ea-4cb9-4089-9b4b-dcb21d83123b.png)

The only stock to experience a negative return in 2017 was TerraForm Power Operating (TERP), and while this stock dropped in value, the 7.2% decline is relatively minor compared to some of the substantial drops seen in the 2018 table. However, the fact that TERP was unable to capitalize on an apparent 2017 boom in the alternative energy sector suggests that it might not be the safest or smartest company to invest in going forward.

On the other end of the spectrum, the only two companies to reach a positive return in 2018 were Enphase Energy (81.9%) and Sunrun (84%). These companies increasing their stock price while other firms in their sector experienced declines as high as 62.6% suggests that they may be capable of resisting negative shocks to the alternative energy sector in the future.

There isn't an obvious conclusion to be reached when observing the total daily volume of the 12 stocks in 2017 and 2018. 7 of the stocks saw higher total volume in 2018 compared to 2017 while 5 of the stocks were traded less frequently. Given that the total volume of the overall sector remained relatively steady, I can conclude that the sector didn't change very much from a volatility standpoint from year to year.

In attempt to improve the run time of the stock analysis, I refactored my original code. Before refactoring the code, the algorithm was such that for every new ticker, the subroutine scanned every single row of the stocks data (either the 2017 or 2018 sheet). After refactoring, the subroutine was made more efficient to require just one total scan through the data. Instead of going back to the first row of data when moving to the next ticker, the refactored subroutine holds the current row and updates the ticker index to the next integer. More specifically, within the original code, the ticker is updated and then a for loop is executed to check every row starting from the first row of data. The below image shows this process.


![Original Code screenshot](https://user-images.githubusercontent.com/95651156/149552148-a03d5ec7-cf5c-47bb-954d-e7f0c3cfe2e5.png)

In the refactored code, the ticker index is instead updated within the for loop that runs through each row of the data. This process is showed in the below code. 


![Refactored Code Screenshot](https://user-images.githubusercontent.com/95651156/149552300-9bf25168-94be-4783-b28c-6024d898e470.png)

This change, while subtle, lowered the run time of the subroutine substantially:

Run Times with Original Code

![Slow_2017](https://user-images.githubusercontent.com/95651156/149552557-42acb42b-d98d-4809-9467-5ebe66b6a5ac.png)
![slow 2018](https://user-images.githubusercontent.com/95651156/149552571-2f885e70-96de-45bc-b73a-b4452182beb6.png)

Run Times with Refactored Code

![VBA_Challenge_2017](https://user-images.githubusercontent.com/95651156/149552729-d0f122cf-e47c-4672-846a-c96718f88ba9.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/95651156/149552744-8ef72215-71a0-42b1-b470-a162897138d9.png)

The above images show that the refactored code ran in approximately 1/5 the amount of time as the original code.

##Summary

There are both advantages and disadvantages to refactoring code. The obvious advantage, as discussed in the previous section, is that refactoring code can lower the run time of the subroutine. This is especially important when the subroutine has a significant run time to begin with. If the runtime was originally multiple minutes, for example, refactoring the code could lead to saving minutes, if not hours, when running the subroutine in the long run. However, this advantage is far less meaningful when talking about subroutines such as my stock analysis that already ran in less than one second. The time that refactoring the code saved is extremely marginal considering that less than one second is saved every time the subroutine is ran.

The main disadvantage of refactoring code is that it requires the coder to put more time and energy into creating the subroutine without changing the algorithm's overall output. If refactoring the code is unlikely to save much time when the subroutine is running later, the coder should consider if it's worth it to take the time to refactor. In my case, refactoring the code for the 12-stock analysis required far more time than it saved me. However, if I was trying to run the subroutine on a data set that held significantly more stocks, refactoring would likely be worthwhile because of the substantial lowering of the run time.



