# VBA Challenge

![](Images/stockmarket0.jpg)

###### by Kurt Pessa

----------


#### Setup

1. Created a new repository `VBA-challenge` on GitHub with share-able link at [https://github.com/kpessa/VBA-challenge](https://github.com/kpessa/VBA-challenge)
2. Created a folder to correspond to the challenge called `VBAStocks`


#### Section 1: For Loop

- Created a script that will loop through all the stocks for one year ..

	![](Images/forloop1.png)
	
	-------------

	![](Images/startrow0.png)
	
	-------------
	 
	![](Images/endrow1.png)
	![](Images/endrow0.png)
	-------------

#### Section 2: StockClass Class Module

-  kept track and outputted the following information to a summary table by creating a custom vba class called `Stock Class`
	1. `tickerSymbol` - the ticker symbol
	2. `yearlyChange` - yearly change from opening price at the beginning of the year to the closing price at the end of that year
	3. `percentChange` - the percent change from opening price at the beginning of a given year to the closing price at the end of that year
	4. `totalStock` - the total stock volume of the stock

	![](Images/customStockClass.png)

#### Section 3: Traversing through data logic

![](Images/looplogic.png)

#### Section 4: Quality Assurance

![](Images/qa.png)

#### Section 5: Enhancing Performance

- Original macro took about 18.9 seconds to traverse through the 797,711 rows of stock data.

![](Images/performance1.png)	

- Took advice from "Excel Macro Mastery" .. [How to make your Excel VBA code run 1000 times faster.](https://www.youtube.com/watch?v=GCSF5tq7pZ0) 
- Decided to load data into an array before looping through. 
	- Brought processing time down from **~18.9 s to ~2.08 s. ~9 times faster.**  

![](Images/performance2.png)
![](Images/looplogicarray.png)
 
- Also, added a few performance enhancement tricks.

![](Images/performance3.png)