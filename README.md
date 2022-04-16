# VBA Stock Analysis

## Overview of Project
	
	This project uses a dataset of 12 stocks and their daily open price, close price, and volume for the years of 2017 and 2018.
	VBA code is used to perform quick analysis of each stock for the desired year. The stocks' daily performances are summarized into a yearly performance report.
	
### Purpose

	The purpose of this project is to have a display of stocks' total yearly volume and total yearly return percentage formatted to quickly analysis if a stock performed well to help determine if it is a good investment. 	

## Analysis and Challenges


### Analysis of Stocks in 2017
![alt text](https://github.com/zimmer3-iii/VBA_StockAnalysis/blob/main/Resources/VBA_Challenge_2017.png)

	Running the analysis report for the stocks in 2017, we can easily report a few things:
	
		1.) Only TERP had a negative yearly return.
		2.) Since 11 of the 12 stocks had positive yearly returns, the market was probably a good performing year.
			a.) About half the stocks where in the 5 to 50% returns and the other half was above 100%. 
		3.) There doesn't appear to be a correlation between total volume and total return. 

### Analysis of Stocks in 2018
![alt text](https://github.com/zimmer3-iii/VBA_StockAnalysis/blob/main/Resources/VBA_Challenge_2018.png)
	
	Running the analysis report for the stocks in 2018, we can easily report a few things:
	
		1.) Only 2 stocks had a positive return: ENPH & RUN 
			a.) Safe to say the stock market had a poor performing year.
			b.) RUN is the highest return in 2018 but was one of the lowest in 2017.
			c.) ENPH was the second highest return in 2018 and was the third highest in 2017, appears to be a great investment.
		2.) Like in 2017, there doesn't appear to be any correlation between total volume and total return.
	

### Challenges and Difficulties Encountered

		1.) The biggest and only challenge I had was do to using starter code and try to refactor it.
	The starter code was incomplete. Trying to follow someone elses thought process to finish the VBA code was difficult. Will writing the code from
	scratch may have taken a little longer, I feel it would have been easier and cleaner. 

## Results

	While the result of this code was not different from previous code used to analyze the same data, the code did run faster by a factor of 10 at least. 
	Refactoring the code took more effort to think through what had already been done, what needed to be added, and what can/should be removed. 

	The main differences in how the code runs compared to the previous code use is that there is:
		1.) There are no nested For loops.
		2.) The data is collected and stored for all stocks before displaying it in the cells.
			a.) Previous code collected data for 1 stock and then displayed it before moving onto the next stock.
			
	While I can't say for certain, the two changes mentioned above seem to be the reason the code runs so much faster. Maybe switch between sheets to collect and display data with the old code was what caused the delay. Having a nested For loop doesn't seem like it would slow anything down, but maybe it is how it was being used to collect and display one stock at a time.

