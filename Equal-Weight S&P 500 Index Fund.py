import pandas as pd
import numpy as np
import yfinance as yf
import math
import xlsxwriter

# Load list of stocks
filepath = '/Users/ahamedrashid/Desktop/Projects/Trading Algorithm Project/Strategy Test/Follow Supplement/sp_500_stocks.csv'
list_stocks = pd.read_csv(filepath)

# Add Stock Data into Pandas DataFrame
new_column = ['Ticker', 'Price', 'Market Capitalization', 'Number of Shares to Buy']
new_dataframe = pd.DataFrame(columns=new_column)  # Create DataFrame with specified columns

# Looping tickers in the List of Stocks and output data to DataFrame
rows = []
for symbol in list_stocks['Ticker']:  # Assuming the CSV has a column named 'Ticker'
    stock = yf.Ticker(symbol)
    
    try:
        market_cap = stock.info.get('marketCap')
        price = stock.info.get('previousClose')
        count = 0
        # Only add to DataFrame if both price and market_cap are available
        if price is not None and market_cap is not None:
            rows.append([symbol, price, market_cap, 'N/A'])
        else:
            print(f"Data not available for ticker: {symbol}")
            count+=1
    except Exception as e:
        print(f"Error fetching data for {symbol}: {e}")
print(count)
# Create the DataFrame all at once
new_dataframe = pd.DataFrame(rows, columns=new_column)

# Number of shares to buy calculation
portfolio_size = float(input('Enter the size of your portfolio: '))
position_size = float(portfolio_size) / len(new_dataframe.index)

for i in range(len(new_dataframe['Ticker'])):
    new_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size / new_dataframe['Price'][i])

# Define the path where you want to save the file
output_filepath = '/Users/ahamedrashid/Desktop/Projects/Trading Algorithm Project/Strategy Test/Equal-Weight S&P 500 Index Fund/recommended_trades.xlsx'

# Create Excel file
with pd.ExcelWriter(output_filepath, engine='xlsxwriter') as writer:
    new_dataframe.to_excel(writer, sheet_name='Recommended Trades', index=False)

    # Excel Formatting
    background_color = '#0a0a23'
    font_color = '#ffffff'

    string_format = writer.book.add_format({
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    })

    dollar_format = writer.book.add_format({
        'num_format': '$0.00',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    })

    integer_format = writer.book.add_format({
        'num_format': '0',
        'font_color': font_color,
        'bg_color': background_color,
        'border': 1
    })

    # Column Formatting
    column_formats = {
        'A': ['Ticker', string_format],
        'B': ['Price', dollar_format],
        'C': ['Market Capitalization', dollar_format],
        'D': ['Number of Shares to Buy', integer_format]
    }

    for column in column_formats.keys():
        writer.sheets['Recommended Trades'].set_column(f'{column}:{column}', 20, column_formats[column][1])
        writer.sheets['Recommended Trades'].write(f'{column}1', column_formats[column][0], string_format)

print(f"Excel file saved to {output_filepath}")
