# Performance Dashboard Ticker

A Google Apps Script project that creates a real-time performance dashboard for tracking financial indices and stocks from multiple markets. This dashboard calculates and displays various return metrics and supports data retrieval from different sources with advanced date handling across time zones.

## Overview

This project is built to run in Google Sheets and provides the following features:

- Track financial indices and stocks from multiple sources (Google Finance, Yahoo Finance, Naver Finance)
- Calculate and display performance metrics:
  - Weekly returns
  - Monthly returns
  - Year-to-date (YTD) returns
  - Returns compared to 52-week high
- Intelligent market awareness with time zone and trading hour support
- Race condition prevention with script locking
- Comprehensive error handling and user feedback
- User-friendly UI with detailed help documentation
- Diagnostic tools for troubleshooting

## Current File Structure

The codebase has been refactored and is now organized into a more maintainable structure:

- **main.js**: Entry point and initialization
- **Config.js**: Configuration settings, constants, and utility functions
- **ui/**
  - **Dialogs.js**: UI dialog-related functionality
  - **Sidebar.js**: Sidebar-related functionality
  - **Sidebar.html**: HTML interface for the sidebar
- **services/**
  - **DashboardService.js**: Core dashboard functionality
  - **TickerService.js**: Ticker-related functionality
  - **PriceService.js**: Price retrieval and calculation
  - **providers/**
    - **BaseProvider.js**: Abstract base class for data providers
    - **ProviderFactory.js**: Factory for creating data providers
    - **GoogleProvider.js**: Google Finance data retrieval
    - **YahooProvider.js**: Yahoo Finance data retrieval
    - **NaverProvider.js**: Naver Finance data retrieval
- **utils/**
  - **DateUtils.js**: Date manipulation and market time management
  - **SpreadsheetUtils.js**: Spreadsheet utility functions
  - **LockService.js**: Script locking functionality
  - **Logger.js**: Enhanced logging functionality
- **diagnostics/**
  - **DiagnosticTools.js**: Diagnostic functionality
- **__tests__/**
  - **integration/**: Integration tests
    - **naverKosdaqPlaywright.test.js**: PlayWright test for KOSDAQ data
    - **naverKosdaqFetcher.test.js**: API fetcher test for KOSDAQ data
  - **mocks/**: Test mocks
  - **services/**: Service unit tests
- **appsscript.json**: Project configuration file

## Main Components

### Core Classes

1. **DateCalculator**: Handles date calculations for historical returns
2. **MarketTimeManager**: Manages market-specific dates and trading hours
3. **DataProviderFactory**: Creates appropriate data providers based on the source
4. **GoogleFinanceProvider/YahooFinanceProvider/NaverFinanceProvider**: Source-specific data retrieval

### Key Functions

1. **updatePerformanceDashboard()**: Main entry point for updating the dashboard
2. **processTickers()**: Processes each ticker for display
3. **calculateReturns()**: Calculates performance metrics
4. **getReferenceDate()**: Gets the reference date for calculations
5. **showReferenceDateDialog()**: UI for setting the reference date
6. **showTickerManager()**: UI for managing tickers
7. **showHelp()**: Displays help documentation

### UI Entry Points

1. Menu items from **onOpen()** function:
   - Dashboard Update
   - Reference Date Settings
   - Ticker Management
   - Diagnostic Tools
   - Help

## Current Status

The codebase has been successfully refactored from a monolithic structure into a more maintainable organization with proper separation of concerns:

- Code is now organized by feature and responsibility in smaller, focused modules
- Data providers have been enhanced with improved error handling and fallback mechanisms
- Integration tests help ensure data retrieval reliability
- Main components now have consistent interfaces and naming conventions
- Provider factory pattern allows for easier extension with new data sources

### Ongoing Improvements

While major refactoring has been completed, these areas continue to be enhanced:

1. **Testing**: Adding more unit tests for individual services and utilities
2. **Error Handling**: Further improving error handling and user feedback for edge cases
3. **Documentation**: Enhancing inline documentation for complex methods
4. **Performance**: Fine-tuning request patterns to reduce API calls

## Implementation Highlights

The refactored implementation includes these key improvements:

1. **Provider System**: Flexible provider system with a factory pattern for extensibility
2. **Advanced KOSDAQ Handling**: Specialized methods for Korean market indices like KOSDAQ
3. **Multi-tier Fallbacks**: Every provider includes multiple fallback mechanisms for reliability
4. **Enhanced Date Handling**: Sophisticated date calculations with market-aware adjustments
5. **Integration Testing**: PlayWright and API-based tests to validate end-to-end functionality

## Setup and Usage

1. Create a Google Sheet with the following structure:
   - "dashboard" sheet: For displaying the performance dashboard
   - "tickers" sheet: For configuring the indices/stocks to track

2. In the "tickers" sheet, add columns:
   - Name: Display name for the index/stock
   - Ticker: Symbol used by the data source
   - Source: One of: "google", "yahoo", or "naver"

3. Use the menu items to manage the dashboard:
   - Update the dashboard to retrieve the latest data
   - Set a reference date for calculations
   - Manage tickers for tracking
   - View help documentation

## Data Sources

The script supports three data sources:

- **Google Finance**: Uses GOOGLEFINANCE function
- **Yahoo Finance**: Fetches data via Yahoo Finance APIs
- **Naver Finance**: Fetches data from Naver Finance (Korean markets)

## Development

This project is managed using [clasp](https://github.com/google/clasp), Google's Command Line Apps Script Projects tool.

To work with this code locally:

1. Install clasp: `npm install -g @google/clasp`
2. Clone the project: `clasp clone <scriptId>`
3. Make changes locally
4. Push changes: `clasp push`

## Running Integration Tests

The project includes integration tests to verify data provider functionality, especially for challenging sources like KOSDAQ data. To run these tests:

### PlayWright Tests

1. Install PlayWright CLI: `npm install -g playwright`
2. Run the KOSDAQ test: `node __tests__/integration/naverKosdaqPlaywright.test.js`

These tests are valuable for diagnosing issues when the data providers stop working due to source website changes. They help understand the correct API endpoints and response formats to implement in the provider classes.

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

### May 2025 - KOSDAQ Data Retrieval Improvements

The following improvements have been made to the KOSDAQ index data retrieval:

1. **Naver Finance Provider**:
   - Implemented specialized method for KOSDAQ historical data retrieval based on Playwright test insights
   - Updated the API endpoint format from `?pageSize=${apiPageSize}&page=1&type=index` to `?timeframe=${timeframe}`
   - Added proper timeframe parameters (1d, 3m, 6m, 1y) based on the historical data range needed
   - Added additional fallback mechanisms for data access:
     - Chart API endpoint as a secondary data source
     - Improved estimation methods when API data is unavailable
   - Enhanced logging for better debugging of API responses

2. **Integration Testing**:
   - Added Playwright-based integration test `naverKosdaqPlaywright.test.js` to validate KOSDAQ historical data fetching
   - This test directly accesses Naver's mobile site to confirm API endpoint structures and response formats
   - Test validates weekly, monthly, and YTD prices against current values

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