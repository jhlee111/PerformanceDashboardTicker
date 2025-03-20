# Performance Dashboard Ticker

A Google Apps Script project that creates a performance dashboard for tracking various financial indices, stocks, and their returns over different time periods.

## Overview

This project is built to run in Google Sheets and provides functionality to:

- Track multiple financial indices and stocks from various sources (Google Finance, Yahoo Finance, Naver Finance)
- Calculate and display performance metrics:
  - Weekly returns
  - Monthly returns
  - Year-to-date (YTD) returns
  - Returns compared to 52-week high

## File Structure

The codebase is organized into multiple files for better separation of concerns:

- **Config.js**: Contains configuration constants and global variables
- **Code.js**: Main application logic and entry points
- **DateUtils.js**: Date calculation utilities
- **DataProviders.js**: Data provider classes for different finance sources
- **tickers.js**: Functions for retrieving and parsing ticker data

## Setup

1. Open your Google Sheet
2. Set up two sheets:
   - Main sheet: For displaying the performance dashboard
   - "tickers" sheet: For configuring the indices/stocks to track

3. In the "tickers" sheet, create columns for:
   - Name (Display name for the index/stock)
   - Ticker (Symbol used by the data source)
   - Source (One of: "google", "yahoo", or "naver")

4. In cell A1 of the main sheet, enter a reference date for calculations

## Usage

Run the `createStockTable()` function to generate the performance table. The script will:

1. Read the reference date from cell A1
2. Get the list of tickers from the "tickers" sheet
3. Fetch current and historical prices for each ticker
4. Calculate performance metrics
5. Render the results in the main sheet

## Data Sources

The script supports three data sources:

- Google Finance: Uses the GOOGLEFINANCE function
- Yahoo Finance: Fetches data via Yahoo Finance API
- Naver Finance: Scrapes data from Naver Finance website (Korean market indices)

## Development

This project is managed using [clasp](https://github.com/google/clasp), Google's Command Line Apps Script Projects tool.

To work with this code locally:

1. Install clasp: `npm install -g @google/clasp`
2. Clone the project: `clasp clone <scriptId>`
3. Make changes locally
4. Push changes: `clasp push`

## File Execution Order in Google Apps Script

In Google Apps Script, files are executed in the order they appear in the Apps Script editor. For this project, the correct order is:

1. Config.js
2. DateUtils.js
3. DataProviders.js
4. tickers.js
5. Code.js

This ensures that dependencies are properly loaded before they're used.

## Functions

- `createStockTable()`: Main function to generate the performance table
- `getTickerList()`: Gets the list of tickers from the "tickers" sheet
- `getStockPrice()`: Gets the price for a given ticker from the specified source
- `calculateReturn()`: Calculates percentage return between two prices

## Recent Updates

### March 2025 - Data Provider Fixes

The following improvements have been made to the data providers:

1. **Naver Finance Provider**:
   - Updated HTML parsing patterns for stock and index prices
   - Added multiple fallback regex patterns to handle changes in Naver Finance website structure
   - Improved extraction of 52-week high values

2. **Yahoo Finance Provider**:
   - Updated API endpoint handling to work with recent Yahoo Finance API changes
   - Added error handling for API response errors
   - Improved extraction of current, historical, and 52-week high prices
   - Added fallback methods for price retrieval when primary methods fail

3. **Testing**:
   - Added a `testDataProviders()` function to verify data provider functionality
   - Run this function to test if the data providers are working correctly

## License

MIT 