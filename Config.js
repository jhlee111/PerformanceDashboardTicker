/**
 * Performance Dashboard Ticker - Configuration Module
 * 
 * Contains all configuration settings for the application.
 */

/**
 * Global configuration object
 */
var CONFIG = {
  /**
   * Sheet names and references
   */
  SHEETS: {
    DASHBOARD: "Dashboard",  // Main dashboard display sheet
    TICKERS: "Tickers",      // Sheet containing ticker definitions
    TEMP: "TempCalc",        // Hidden sheet for temporary calculations
    AUDIT: "Audit"           // Added for audit purposes
  },
  
  /**
   * Display settings
   */
  DISPLAY: {
    DASHBOARD_SHEET: "dashboard",
    DATA_RANGE_START: "A3",
    HEADER_ROW: 2,
    DATA_RANGE_END: "F"      // Ending column for data display
  },
  
  /**
   * Dashboard specific settings
   */
  DASHBOARD: {
    INFO_ROW: 1,             // Row for info/status messages
    INFO_COL: 1,             // Column for info/status messages
    HEADER_ROW: 2,           // Header row index
    DATA_START_ROW: 3,       // Starting row for data display
    TICKER_COL: 1,           // Column for ticker symbols
    NAME_COL: 2,             // Column for ticker names
    CURRENT_PRICE_COL: 3,    // Column for current prices
    WEEKLY_COL: 4,           // Column for weekly returns
    MONTHLY_COL: 5,          // Column for monthly returns
    YTD_COL: 6               // Column for year-to-date returns
  },
  
  /**
   * Status indicators
   */
  STATUS: {
    NO_DATA: "데이터 없음",     // No data available indicator
    ERROR: "오류",             // General error indicator
    CALC_ERROR: "계산 오류",    // Calculation error indicator
    WARNING: "주의"            // Warning indicator
  },
  
  /**
   * Temporary cells used for formula calculations
   */
  TEMP_CELLS: {
    GOOGLE_PRICE: "A1:A1",   // Cell for Google Finance price formulas
    GOOGLE_HIGH: "B1:B1",    // Cell for Google Finance 52-week high formulas
    YTD_DATE: "C1:C1",       // Cell for YTD date calculations
    GENERAL: "D1:D5"         // General purpose temp cells
  },
  
  /**
   * Date ranges for return calculations (in days)
   */
  DATE_RANGES: {
    WEEKLY: 7,              // One week period
    MONTHLY: 30             // One month period
  },
  
  /**
   * Retry settings for data fetching
   */
  RETRIES: {
    MAX_ATTEMPTS: 10,       // Maximum number of retry attempts
    DELAY_MS: 500           // Delay between retries in milliseconds
  },
  
  /**
   * Property keys for script properties
   */
  PROPS: {
    REF_DATE: "referenceDate"  // Key for reference date in script properties
  }
};

// Get a reference to the active spreadsheet if we don't already have one
if (typeof SS === 'undefined') {
  var SS = SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Get the reference date value safely from script properties
 * @return {Date|null} The reference date or null if not available
 */
function getRefDateValue() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const dateString = scriptProps.getProperty(CONFIG.PROPS.REF_DATE);
    
    if (!dateString) {
      return null;
    }
    
    return new Date(dateString);
  } catch (e) {
    Logger.log(`참조 날짜 값을 가져오는 중 오류가 발생했습니다: ${e.message}`);
    return null;
  }
}

/**
 * Set the reference date value safely in script properties
 * @param {Date|string} value - The date value to set
 * @return {boolean} True if successful, false otherwise
 */
function setRefDateValue(value) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    let dateString;
    
    if (value instanceof Date) {
      // Format as ISO string for storage
      dateString = Utilities.formatDate(value, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
    } else {
      // If it's already a string, parse it first to validate
      const date = new Date(value);
      if (isNaN(date.getTime())) {
        return false;
      }
      dateString = Utilities.formatDate(date, "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
    }
    
    scriptProps.setProperty(CONFIG.PROPS.REF_DATE, dateString);
    return true;
  } catch (e) {
    Logger.log(`참조 날짜 값을 설정하는 중 오류가 발생했습니다: ${e.message}`);
    return false;
  }
}

/**
 * Safely get a configuration value with fallback
 * @param {string} path - Dot notation path to config value
 * @param {any} defaultValue - Fallback value if path not found
 * @return {any} Config value or default
 */
function getConfigValue(path, defaultValue) {
  try {
    const parts = path.split('.');
    let current = CONFIG;
    
    for (const part of parts) {
      if (current[part] === undefined) {
        return defaultValue;
      }
      current = current[part];
    }
    
    return current;
  } catch (e) {
    return defaultValue;
  }
} 