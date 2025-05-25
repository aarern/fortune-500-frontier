# Fortune 500 Frontier

The graph creation was sourced from this Github user: https://github.com/tthustla/efficient_frontier/blob/master/Efficient%20_Frontier_implementation.ipynb
The web scraper, ticker, and yield importing, and VBA scripts were my contributions.

Description:
This script allows you to make a custom portfolio with Fortune 500 companies and optimize your returns with the Markowitz efficient frontier. This model pulls in the last 30 months' worth of daily stock prices, and the treasury yields being pulled in are on a floating basis to properly calculate your returns in excess of the risk-free rate. This workbook also contains preset industries and portfolios so you can experiment with the industry returns. 

When the graph populates, the blue star will represent your optimal portfolio that is filtered for outliers (ex, a portfolio that has a 90% investment in NVDA). All optimal portfolios (minimum volatility, max sharpie, max sharpie filtered) will print the weights of each stock in the terminal.

I included a web scraper to practice importing data from the internet, but it is not necessary for this exercise. The macro-enabled workbook also isn't necessary for this exercise, but the automated formatting and portfolio customization make it much easier to use with the Python script.

If you want a better representation of the SPY portfolio, increase the num_portfolios integer in the script so you have a better representation of the most efficient portfolio. I ran this script with over 500k simulations in one test and it creates a much clearer picture. 

How to get to work:
1. Open Excel
2. Trust Center settings 
3. Trusted locations
4. If using a Linux terminal, "Allow Trusted Locations on my Network (not recommended)" enable
5. Update the "file_path" to the location of the macro-enabled workbook
6. Update the "num_portfolios" to your desired size. The higher the portfolio simulations, the more accurate the portfolio is

pip installation list:
1. pip install pandas 
2. pip install requests
3. pip install requests beautifulsoup4
4. pip install openpyxl
5. pip install yfinance
5. pip install --upgrade yfinance
6. pip install pandas_datareader
7. pip install tqdm
8. pip install python3
