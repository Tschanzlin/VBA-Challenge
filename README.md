# VBA-Challenge Notes:

- Created a VBA script that ran and correctly calculated the following:
-- For each ticker:
	> Dollar price change
	> Percentage price change
	> Total volume traded
-- For each worksheet:
	> Greatest % increase and associated ticker
	> Greatest % increase and associated ticker
	> Greatest total volume

- The VBA loops correctly ran through each worksheet for Alphatesting.  I have attached screens shot of alpha testing tabs A and B to show that the VBA program correctly calculated and formatted the values by ticker and worksheet.  

- The VBA loops have not consistently run throgh the Multiple-year workbook.  It did work on more than one occasion, but now is timing out on the first spreadsheet.  If you run the script, you'll be able to see the values correctly calculate the maximum and minimum price increases and volume for any individual "year" worksheet, but the loop hangs up.

- I have tried multiple trouble shooting efforts to fix the issue, including:
	-- Reducing number of spreasheets to loop from three to two
	-- Eliminating worksheet loop altogether
	-- Hardcoding "LastRow" value so that LastRow is not calculated

- The debugging effort has highlighted two areas -- the MinTicker variable or Next i command.  I would continue to trouble shoot these two points or possibly change the looping format to a Do While loop that would continue to execute while the cell value was greater than zero. 

- I've also wondered whether data error in the larger excel file or if my computer does not have enough memory or processing power to work on the larger datafile.  I can't otherwise explain why the program works in one workbook and not the other.


	

	