# stock-analysis
This script is designed to analyze stock data for one year and output various information including the ticker symbol, yearly change, percentage change, and total stock volume. Additionally, it includes functionalities to identify the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume. The script is then applied to each sheet to give a multi-year analysis.

## Process
To achieve the desired results, I first focused on Retrieval of Data. This involved looping through one year of stock data and retrieving and storing the following values from each row: Ticker symbol, Volume of stock, Open price, and Close price.
Next, I proceeded with Column Creation, ensuring all necessary columns were created on the same worksheet as the raw data or on a new worksheet. This included correctly creating columns for
Ticker symbol, Total stock volume, Yearly change ($), and Percent change.
Following column creation, I applied Conditional Formatting to the yearly change column and ensured appropriate application to the percent change column.

<img width="769" alt="Screenshot 2024-03-26 at 1 13 09 PM" src="https://github.com/mattflanagan2/stock-analysis/assets/146908072/4c0a8337-8d6e-4dba-95b0-5f35d80af558">
<img width="819" alt="Screenshot 2024-03-26 at 1 13 33 PM" src="https://github.com/mattflanagan2/stock-analysis/assets/146908072/0286d7de-4290-4f16-86e5-ac3f9dc8e49d">




 Moving on to Calculated Values, I calculated and displayed the Greatest % Increase, Greatest % Decrease, and Greatest Total Volume. Lastly, I ensured the VBA script could run successfully on all sheets, ensuring seamless looping across worksheets.

<img width="810" alt="Screenshot 2024-03-26 at 1 19 09 PM" src="https://github.com/mattflanagan2/stock-analysis/assets/146908072/6f76bd79-9adf-49c1-b363-b835d0c88118">
<img width="810" alt="Screenshot 2024-03-26 at 1 20 02 PM" src="https://github.com/mattflanagan2/stock-analysis/assets/146908072/f3f1a38d-f17f-4082-8ca7-458e1c6b0739">
<img width="810" alt="Screenshot 2024-03-26 at 1 20 22 PM" src="https://github.com/mattflanagan2/stock-analysis/assets/146908072/8cf0d8a7-7ffe-4cf6-b177-1cc9db9b7abc">
  
  
  
  

## Final Results

### Stock Analysis for 2018
![2018](https://github.com/mattflanagan2/stock-analysis/assets/146908072/3b38534f-aa9e-485d-8d35-636752c73af8)

### Stock Analysis for 2019
![2019](https://github.com/mattflanagan2/stock-analysis/assets/146908072/2bf4e62c-34d1-44af-87d4-f23ef1a41618)


### Stock Analysis for 2020
![2020](https://github.com/mattflanagan2/stock-analysis/assets/146908072/48c407ae-b393-4a4b-8e73-7fffcf5dc9c1)





