# VBA-Challenge
Stock Market Analysis with VBA

## Contents
1. [Overview](#1-overview)  
2. [Repository](#2-repository)  
3. [Deployment](#3-deployment)  
4. [Data Analysis](#4-data-analysis)  
5. [References](#5-references)  


## 1. Overview
This challenge focuses on analysing stock market data using VBA (Visual Basic for Applications) scripting in Excel. The task is to create a script that loops through quarterly stock data and calculates key metrics, including the ticker symbol, quarterly price changes, percentage change, total stock volume, and highlights the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. Additionally, the script applies conditional formatting to highlight positive and negative changes in stock prices, making the results visually clear and easy to interpret.

## 2. Repository
The repository contains the following files:

- `StockAnalysis.vbs` - The VBA script used for stock data analysis.
- [Resources/](Resources) - Folder containing the Excel file that the VBA script applies to.
  - `Multiple_year_stock_data.xlsx` - Excel file (warning: large file size)
- [`Screenshots/`](Screenshots) - Folder containing screenshots of the results for different years:
  - [`Screenshot_2018.png`](Screenshots/Screenshot_2018.png)
  - [`Screenshot_2019.png`](Screenshots/Screenshot_2019.png)
  - [`Screenshot_2020.png`](Screenshots/Screenshot_2020.png)


## 3. Deployment
To run the analysis, follow these simple steps:
1. Download this repository to your local computer.
2. Import the `StockAnalysis.vbs` script into the Excel file `Multiple_year_stock_data.xlsx`.
3. Run the module `StockAnalysis()` to execute the script.  
   The script will process all available data for each quarter and output the required analysis in the worksheet.


## 4. Data Analysis
The script analyses stock data for each quarter and calculates the following:
- **Ticker Symbol**: Identifies the stock being analysed.
- **Quarterly Change**: The difference between the opening price at the start of the quarter and the closing price at the end of the quarter.
- **Percentage Change**: The percentage change in the stock's price over the quarter.
- **Total Stock Volume**: The total volume of the stock traded during the quarter.

The script also identifies and returns:
- The **Greatest % Increase** in stock price across all quarters.
- The **Greatest % Decrease** in stock price across all quarters.
- The **Greatest Total Volume** of stocks traded in a quarter.

Results are displayed directly in the worksheet with the appropriate columns, and conditional formatting is applied to visually highlight:
- Positive changes in stock prices in **green**.
- Negative changes in stock prices in **red**.


## 5. References
- Data for this dataset was generated by **edX Boot Camps LLC** and is intended for educational purposes only.
