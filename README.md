# VBA Stock Market Analysis

<img src="https://assets.bwbx.io/images/users/iqjWHBFdfxIU/iQihAS88YxEU/v3/620x-1.jpg" alt="Crowdfunding Project Analysis" style="width: 100%; height: auto;">

***

## Overview

This project uses VBA scripting to analyze stock market data across multiple quarters. The goal is to calculate and report various metrics for each stock, including quarterly changes, percentage changes, and total volumes. The script also identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume. 

***

## Features

- **Ticker Symbol Analysis**: Extracts and processes ticker symbols for each stock.
- **Quarterly Change Calculation**: Computes the quarterly change in stock prices.
- **Percentage Change Calculation**: Computes the percentage change from the opening price to the closing price of each quarter.
- **Total Volume Calculation**: Sums up the total volume of stock traded.
- **Greatest Values Identification**: Identifies the stocks with the greatest percentage increase, decrease, and total volume.
- **Conditional Formatting**: Applies conditional formatting to highlight positive changes in green and negative changes in red.

***

## Instructions

1. **Setup the Repository**:
    - Create a new repository named `VBA-challenge` on GitHub.
    - Clone the repository to your local machine.

2. **Add VBA Script**:
    - Open the Excel workbook you wish to analyze.
    - Press `ALT + F11` to open the VBA editor.
    - Insert a new module and paste the provided VBA script (`StockMarketAnalysis.vba`).

3. **Run the Script**:
    - Run the `StockMarketAnalysis` macro.
    - The script will process the data in each sheet, calculate the required metrics, and output the results.

4. **Review the Results**:
    - The results, including the ticker symbol, total volume, quarterly change, and percentage change, will be displayed in the new columns.
    - The greatest percentage increase, greatest percentage decrease, and greatest total volume will be highlighted separately.

***

## Example Results

| Ticker Symbol                                                                       | Calculations                                                                   |
|-------------------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| ![Moderate Solution](https://static.bc-edx.com/data/dl-1-2/m2/lms/img/moderate_solution.jpg) | ![Hard Solution](https://static.bc-edx.com/data/dl-1-2/m2/lms/img/hard_solution.jpg) |

***

## Code Explanation

snippet 
```vba
' Calculate yearly change and percentage change
yearlyChange = closingPrice - openingPrice
If openingPrice <> 0 Then
    percentChange = (closingPrice - openingPrice) / openingPrice * 100
Else
    percentChange = 0
End If
```

The VBA script performs the following steps:

1. **Initialization**:
    - Declares variables for ticker symbols, prices, volumes, and changes.
    - Initializes variables for tracking the greatest values.

2. **Loop Through Worksheets**:
    - Iterates over each worksheet to process quarterly data.

3. **Loop Through Rows**:
    - Processes each row of data to compute the required metrics.
    - Detects changes in ticker symbols to calculate quarterly changes and percentages.

4. **Output Data**:
    - Stores the calculated data in new columns.
    - Applies conditional formatting to highlight changes.
    - Tracks and outputs the greatest values for percentage change and volume.

5. **Conditional Formatting**:
    - Applies green color for positive changes and red for negative changes.

***

## Key Insights

- **Positive and Negative Changes**: The script highlights positive changes in green and negative changes in red, making it easy to spot trends.
- **Greatest Values**: Identifies and displays the stock with the greatest percentage increase, greatest percentage decrease, and greatest total volume.
- **Automated Analysis**: Automates the tedious process of analyzing stock data across multiple sheets, saving time and reducing errors.

***


