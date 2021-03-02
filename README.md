# Stocks Analysis
Data Analytics Boot Camp Module 2 Challenge
## Project Overview
Steve’s parents are considering investing their money in the stock market, however they are not so familiar with the products. DQ drew their attention by its name, but Steve wants to give a closer analysis on the stocks’ volumes and prices rather than only names. Considering the massiveness of the data in the stocks market, Steve is considering seeking for assistance from using VBA. The project is to help Steve create a workbook with VBA scripts that can analyze a list of stocks in 2017 and 2018 by their daily volumes and annual return rates.

## Results
The analysis results show that the best performers from 2017 to 2018 are ENPH, SEDG and RUN.

The refactored script can run the analysis as fast as 0.0234 second, which is only 17.6% of the time spent by the script before refactored.

### Stocks Performance in 2017 and 2018
2017 was a bullish year for most of the stocks on the list, as shown in the table below (*Table 1*). A third of the stocks doubled their prices in the end of the year. DQ performed the best among all the stocks listed, with 199.4% return rate. However, the total daily volume of DQ was the lowest in the list. 35,796,200 made DQ being not so reliable. TERP performed the worst in 2017 with -7.2% return, being the only stock had negative return.

*Table 1 All Stocks Analsys in 2017*

![VBA_Challenge_2017_table](https://user-images.githubusercontent.com/78275082/109587088-76387900-7ad4-11eb-9b14-cf746ffd39d0.png)

2018, in contrary, was a bearish year for 10 out of 12 stocks listed, as shown in the table below (*Table 2*). Only 2 stocks, ENPH and RUN, successfully increased their value for roughly more than 80%. The trade volumes of these two stocks were over 500 million, which were higher than that of most of other stocks had in 2018. The two worst performing stocks in 2018 was DQ with -62.6% return and JKS with -60.5% return.

*Table 2 All Stocks Analsys in 2018*

![VBA_Challenge_2018_table](https://user-images.githubusercontent.com/78275082/109587105-7fc1e100-7ad4-11eb-8a72-a963ce51ab8c.png)

The top best performers in the 12 stocks list should be ENPH, SEDG and RUN. Being the winner, ENPH accomplished as high as 350% return from 2017 to 2018, with more than 800 million trade volume. The second place SEDG had 166% return, even though it didn’t perform well in 2018. Same as the third place RUN, which had 95% overall return but more total daily volume than SEDG. More data are presented below in *table 3*.

*Table 3 All Stocks Analsys from 2017 to 2018*

![VBA_Challenge_2017+2018_table](https://user-images.githubusercontent.com/78275082/109587112-83556800-7ad4-11eb-9bea-865ac879e10c.png)

### Original Script vs. Refactored Script
The original script completes the analysis on either 2017 or 2018 stocks data in 0.14 second, while the refactored script completes the task with 0.03 second. The analysis is dramatically increased by simplifying calculation steps and methods. Below are the screenshots (*Graph 1*, *Graph 2*, *Graph 3* and *Graph 4*).

![VBA_Challenge_2017_original](https://user-images.githubusercontent.com/78275082/109589685-b26dd880-7ad8-11eb-8403-6dec1ccf3954.png)

(*Graph 1 Original script's elapsed run time in 2017*)

![VBA_Challenge_2018_original](https://user-images.githubusercontent.com/78275082/109589696-b699f600-7ad8-11eb-8598-af79f130577a.png)

(*Graph 2 Original script's elapsed run time in 2018*)

![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/78275082/109589301-2065d000-7ad8-11eb-82bd-cea93f044468.png)

(*Graph 3 Refactored script's elapsed run time in 2017*)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/78275082/109589307-222f9380-7ad8-11eb-9e8c-8d2753ff8c36.png)

(*Graph 4 Refactored script's elapsed run time in 2018*)

The refactored script completes the analysis with single looping through the stocks data. This simplified process is accomplished by establishing a `tickerIndex`, which helps record the stock’s name, volume and prices as the loop goes through. The time is saved by making every loop take the arguments and add some values. On the other hand, the original script creates 11 times more loops that only make arguments without calculating.

The methods used to determine the last row of a worksheet also make a difference. The ` rowEnd = Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row` runs much faster than ` rowEnd = Cells(Rows.Count, "A").End(xlUp).Row`.

A special trick is adding ` Application.ScreenUpdating = False` to the beginning of the script and adding it back in the end as ` Application.ScreenUpdating = True`, if the script does not need to update the screen.

## Summary
Refactoring code can make the subroutine run faster and code lines being shorter than the original codes. Nowadays, with the access to date sets being bigger and bigger, creating efficient codes saves huge amount of time. However, refactoring code too far may not be as beneficial as no matter how fast a code run, the time elapsed cannot be negative. Refactoring may also introduce errors to the code if the steps are not being taken carefully.

On this project specifically, which the current data set is not so big, the code does not necessarily need to be extremely fast. The core element in the refactored script is `tickerIndex`, which could create errors if was used incorrectly.
