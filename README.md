STEPS 
STEP 1. Fetch Live Data
We'll use a public API like CoinGecko API to fetch the data for the top 50 cryptocurrencies by market capitalization.
•	Use Python with the requests library to fetch the data.
•	Choose the fields required for the task, such as Name, Symbol, Current Price (in USD), Market Capitalization, 24-hour Trading Volume, and 24-hour Price Change percentage.

STEP 2. Data Analysis
Perform analysis on the live data to identify key insights:
•	Top 5 Cryptocurrencies by Market Cap: Sort the DataFrame by 'Market Capitalization' and pick the top 5.
•	Average Price of Top 50 Cryptos: Calculate the average of the 'Current Price (USD)' column.
•	Highest and Lowest 24-hour Price Change: Identify the maximum and minimum values in the '24h Price Change (%)' column.

STEP 3. Live-Running Excel Sheet
•	To make the data update live in Excel, you can use Excel's ability to fetch live data directly via Power Query or use a Python package like openpyxl to continuously update the sheet.
•	Alternatively, you can set up a script that updates the data every 5 minutes using time.sleep() in Python.

Final Step
1.	Run the Python script to fetch the live data.
2.	Make the Excel sheet continuously update every 5 minutes.



