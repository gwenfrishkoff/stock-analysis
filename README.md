# An Excel-Based (VBA) Analysis of Kickstarter Campaigns

## Overview of Project
Steve (recent grad with domain knowledge in finance) is our client. Steve's parents are interested in this sector, but have no knowledge or experience in this area and have decided to invest all their money in one business: DAQO New Energy Corp. (ticker symbol 'DQ"), which makes silicon wafers for solar panels. Steve would like to help diversify his parents' investments. To this end, he wishes to analyze DQ stock data and to compare DQ stock with portfolios for other green energy cos. In summary, Steve has two analysis goals: 
	<ol>
	<li> to analyze DQ stock data; and 
	<li> to compare DQ stock with portfolios for other green energy companies in 2017 and 2018.
	</ol>

### Analysis Goals
To support our client aims, we created automated code using VBA (visual basic for applications) to automate excel analysis of green stock data for the years 2017 and 2018. Our code performs the following simple set of analyses:
	<ol>
	<li> It computes the total daily volume for each green stock in each year (2017 and 2018); and
	<li> It computes the yearly return for each green stock in each year (2017 and 2018);
	</ol>

We also provided code to automatically format the results table and created "run" buttons to allow for easy, interactive analyses by the end user.
	
##Results

The resulting data (input and output worksheets) and VBA scripts are provided in the XLSM files ***green_stocks.xlsm*** and ***VBA_Challenge.xlsm***. 
Analysis results are summarized in the figures ***VBA_Challenge_OrigMacro_2017.png*** and ***VBA_Challenge_OrigMacro_2018.png***. 


- Based on the yearly analyses of DQ vs. other stocks, we can draw two general conclusions:
	<ol>
  	<li> In general, DQ stock was outperformed by other green stocks, particularly in 2018. 
  	<li> Two other stocks, SEDG and ENPH, showed the greatest returns in both years.
	</ol>


- Based on the code performance analyses for the original and refactored VBA code, we can draw the following conclusions:
	<ol>
  	<li> Refactoring code can lead to improvements in code performance (e.g., time to execute code);
  	<li> Refactoring code can lead to errors, e.g., in formatting, particularly when there are changes to the data types.
	</ol>
