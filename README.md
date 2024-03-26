# stock-analysis
This script is designed to analyze stock data for one year and output various information including the ticker symbol, yearly change, percentage change, and total stock volume. Additionally, it includes functionalities to identify the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume. The script is then applied to each sheet to give a multi-year analysis.

To achieve the desired results, I first focused on Retrieval of Data. This involved looping through one year of stock data and retrieving and storing the following values from each row: Ticker symbol, Volume of stock, Open price, and Close price. Next, I proceeded with Column Creation, ensuring all necessary columns were created on the same worksheet as the raw data or on a new worksheet. This included correctly creating columns for Ticker symbol, Total stock volume, Yearly change ($), and Percent change. Following column creation, I applied Conditional Formatting to the yearly change column and ensured appropriate application to the percent change column. Moving on to Calculated Values, I calculated and displayed the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume. Lastly, I ensured the VBA script could run successfully on all sheets, ensuring seamless looping across worksheets.








I used the following code to trouble shoot
https://github.com/ermiasgelaye/VBA-challenge/blob/master/VBA_Alphabetical_testing/alphabetical_testing.vbs
  credit to ermiasgelaye

  I also utilized the credit_charges file that we worked on in class to determine my summary table and how to count the tickers
  I also used the census_data_2016-2019_pt1 document 
  
