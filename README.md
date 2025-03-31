# Stock Data Analysis

## Overview
This project automates stock data retrieval and analysis using VBA. The script processes stock market data to extract key metrics, apply calculations, and format results for better visualization.

## Features
- **Data Retrieval**: Extracts ticker symbol, stock volume, open price, and close price.
- **Column Creation**: Computes total stock volume, quarterly change, and percent change.
- **Conditional Formatting**: Highlights key changes in quarterly and percent change columns.
- **Calculated Values**: Identifies the greatest percentage increase, decrease, and total volume.
- **Looping Across Sheets**: Ensures the script runs on all sheets in the workbook.
- **Comprehensive Stock Analysis**: Loops through all stocks to compute yearly and percentage change from the opening and closing prices, total volume for each stock, and stocks with the greatest and least percent change, as well as the greatest total volume.
- **Enhanced Formatting**: The percent change column was formatted using VBA to include the percentage symbol, and green/red color filling was applied to indicate positive or negative values in the yearly change column.

## How to Run
1. Open the Excel file (`Multiple_year_stock_data.xlsm`) containing stock data.
2. Navigate to the **VBA Editor** (`Alt + F11`).
4. Run the script to analyze stock data across all sheets.
5. Review formatted results and calculated insights.

## Analysis Results
### 2018:
- Greatest % Increase: **THB** at 141.42%
- Greatest % Decrease: **RKS** at -90.02%
- Greatest Total Volume: **QKN**

### 2019:
- Greatest % Increase: **RYU** at 190.03%
- Greatest % Decrease: **RKS** at -91.60%
- Greatest Total Volume: **ZQD**

### 2020:
- Greatest % Increase: **YDI** at 188.76%
- Greatest % Decrease: **VNG** at -89.05%
- Greatest Total Volume: **QKN**

### Observations
The most consistent results come from the greatest percent decrease and greatest total volume:
- **RKS** had the greatest percent decrease in both 2018 and 2019.
- **QKN** had the greatest total volume in 2018 and 2020.

This three-year trend suggests that additional data sources should be analyzed to determine correlations between tickers with the highest and lowest percent change, as well as those with the greatest total volume. This analysis could serve as a foundation for using machine learning to predict 2021’s results.

