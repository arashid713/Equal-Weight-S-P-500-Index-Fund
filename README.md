# Equal-Weight SP500 Index-Fund
This Python script analyzes S&P 500 stocks and generates a recommended trades list based on equal weighting of a given portfolio size. The script retrieves stock data using the Yahoo Finance API and outputs the results to an Excel file with custom formatting.
Trading Algorithm Project: S&P 500 Stock Analyzer
This Python script analyzes S&P 500 stocks and generates a recommended trades list based on equal weighting of a given portfolio size. The script retrieves stock data using the Yahoo Finance API and outputs the results to an Excel file with custom formatting.

# Features
Data Retrieval: Fetches stock data including price and market capitalization using yfinance.
Equal Weight Calculation: Distributes a portfolio size equally among the stocks.
Excel Export: Saves the recommended trades to an Excel file with formatted columns.
Requirements
Python 3.x
pandas
numpy
yfinance
xlsxwriter
You can install the required packages using pip:

bash
Copy code
pip install pandas numpy yfinance xlsxwriter

# How to Use
Prepare the Stock List: Ensure you have a CSV file containing a column named 'Ticker' with the list of S&P 500 stock tickers. Update the filepath variable in the script with the path to your CSV file.

Run the Script: Execute the script using Python. It will prompt you to enter the size of your portfolio.

View Results: The script will generate an Excel file with the recommended trades, including stock tickers, prices, market capitalizations, and the number of shares to buy based on the equal-weight strategy.

Check Output: The Excel file will be saved at the specified output_filepath with custom formatting.


# File Paths
Input File: /path/to/your/sp_500_stocks.csv
Output File: /path/to/save/recommended_trades.xlsx

# Notes
Ensure the CSV file has the correct format with a 'Ticker' column.
The script currently supports basic error handling for missing data or errors during data retrieval.
Modify the paths in the script according to your local file structure.
License
This project is licensed under the MIT License - see the LICENSE file for details.

# Acknowledgements
Yahoo Finance API
pandas Documentation
xlsxwriter Documentation
