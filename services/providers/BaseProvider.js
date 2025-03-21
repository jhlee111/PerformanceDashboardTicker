/**
 * Abstract base class for data providers
 */
class DataProvider {
  constructor() {
    this.name = "Default Provider";
    this.diagnosticMode = false;
  }
  
  /**
   * Enable or disable diagnostic mode
   * @param {boolean} enable - Whether to enable diagnostic mode
   */
  setDiagnosticMode(enable) {
    this.diagnosticMode = !!enable;
    Logger.log(`${this.name} diagnostic mode ${this.diagnosticMode ? 'enabled' : 'disabled'}`);
  }
  
  /**
   * Get price for a ticker on a specific date
   * @param {string} symbol - Ticker symbol
   * @param {Date} date - Date to get price for
   * @return {number} Price
   */
  getPrice(symbol, date) {
    throw new Error("getPrice method must be implemented by subclass");
  }
  
  /**
   * Get historical price for a ticker on a specific date
   * @param {string} symbol - Ticker symbol
   * @param {Date} date - Historical date to get price for
   * @return {number} Historical price
   */
  getHistoricalPrice(symbol, date) {
    // By default, delegate to getPrice but subclasses can override with specific implementations
    return this.getPrice(symbol, date);
  }
  
  /**
   * Get the 52-week high price for a ticker
   * @param {string} symbol - Ticker symbol
   * @return {number} 52-week high price
   */
  getHighPrice(symbol) {
    throw new Error("getHighPrice method must be implemented by subclass");
  }
  
  /**
   * Safely extract a value from a potential 2D array result
   * @param {*} result - The result to extract from (could be number, string, or array)
   * @param {number} [rowIndex=1] - Row index to extract (default is 1, the first data row)
   * @param {number} [colIndex=1] - Column index to extract (default is 1, the second column)
   * @return {*} The extracted value or the original result if not an array
   * @protected
   */
  extractValueFromTable(result, rowIndex = 1, colIndex = 1) {
    // If it's not an array or object, return as is
    if (typeof result !== 'object' || result === null) {
      return result;
    }
    
    // If it's an array
    if (Array.isArray(result)) {
      // Log array structure if in diagnostic mode
      if (this.diagnosticMode) {
        this.logArrayStructure(result, "Table result");
      }
      
      // If it's a 2D array with the expected structure
      if (result.length > rowIndex && 
          Array.isArray(result[rowIndex]) && 
          result[rowIndex].length > colIndex) {
        return result[rowIndex][colIndex];
      }
      
      // If it's a 1D array and we're looking for first element
      if (rowIndex === 0 && result.length > 0) {
        return result[0];
      }
    }
    
    // Return the original result if we couldn't extract
    return result;
  }
  
  /**
   * Log the structure of an array for debugging
   * @param {Array} array - The array to log
   * @param {string} name - Name for the log
   * @protected
   */
  logArrayStructure(array, name) {
    if (!array || !Array.isArray(array)) {
      Logger.log(`${name} is not an array`);
      return;
    }
    
    Logger.log(`${name} is an array with ${array.length} rows`);
    
    // Log header row if exists
    if (array.length > 0) {
      if (Array.isArray(array[0])) {
        Logger.log(`Header row: ${JSON.stringify(array[0])}`);
      } else {
        Logger.log(`First element: ${array[0]}`);
      }
    }
    
    // Log data rows (up to 5)
    for (let i = 1; i < Math.min(array.length, 6); i++) {
      if (Array.isArray(array[i])) {
        Logger.log(`Data row ${i}: ${JSON.stringify(array[i])}`);
      } else {
        Logger.log(`Element ${i}: ${array[i]}`);
      }
    }
    
    // If there are more rows, indicate that
    if (array.length > 6) {
      Logger.log(`...and ${array.length - 6} more rows`);
    }
  }
  
  /**
   * Check if a value represents an error
   * @param {*} value - Value to check
   * @return {boolean} True if the value represents an error
   * @protected
   */
  isErrorValue(value) {
    // Check for standard error strings
    if (typeof value === 'string') {
      const errorStrings = ["#N/A", "#REF", "#NAME", "#DIV/0", "#NULL", "#VALUE", "#NUM", "ERROR"];
      for (const errorStr of errorStrings) {
        if (value.includes(errorStr)) {
          return true;
        }
      }
    }
    
    // Check for NaN and other non-numeric values for numeric contexts
    if (this.shouldBeNumeric(value)) {
      return isNaN(value) || value === null || value === undefined;
    }
    
    return false;
  }
  
  /**
   * Determines if a value should be numeric based on context
   * @param {*} value - Value to check
   * @return {boolean} True if the value should be numeric
   * @private
   */
  shouldBeNumeric(value) {
    // Price and percentage values should be numeric
    return !isNaN(parseFloat(value)) || 
           (typeof value === 'string' && value.includes('%'));
  }
  
  /**
   * Formats a value for display
   * @param {*} value - Value to format
   * @param {string} [type="price"] - Type of value (price, percent, date)
   * @return {string} Formatted value
   * @protected
   */
  formatValue(value, type = "price") {
    if (this.isErrorValue(value)) {
      return CONFIG.STATUS.NO_DATA;
    }
    
    try {
      switch (type) {
        case "price":
          return parseFloat(value).toLocaleString(undefined, {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
          });
        
        case "percent":
          const percentValue = parseFloat(value);
          const sign = percentValue > 0 ? "+" : "";
          return `${sign}${percentValue.toFixed(2)}%`;
          
        case "date":
          if (value instanceof Date) {
            return Utilities.formatDate(value, "GMT+9", "yyyy-MM-dd");
          }
          return value.toString();
          
        default:
          return value.toString();
      }
    } catch (error) {
      Logger.log(`Formatting error for ${value}: ${error.message}`);
      return value.toString();
    }
  }
} 