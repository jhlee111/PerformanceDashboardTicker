/**
 * Factory for creating data providers
 */
class DataProviderFactory {
  constructor(marketTimeManager) {
    this.marketTimeManager = marketTimeManager || new MarketTimeManager();
  }
  
  /**
   * Get a data provider for the given source
   * @param {string} source - Data source (google, yahoo, naver)
   * @return {DataProvider} The appropriate data provider
   */
  getProvider(source) {
    if (!source) {
      throw new Error("데이터 소스가 지정되지 않았습니다.");
    }
    
    const lowerSource = source.toLowerCase().trim();
    
    switch (lowerSource) {
      case "google":
        return new GoogleFinanceProvider(this.marketTimeManager);
      case "yahoo":
        return new YahooFinanceProvider(this.marketTimeManager);
      case "naver":
        return new NaverFinanceProvider(this.marketTimeManager);
      default:
        throw new Error(`지원되지 않는 데이터 소스입니다: ${source}`);
    }
  }
  
  /**
   * Format a symbol for a specific provider
   * @param {string} symbol - The ticker symbol
   * @param {string} provider - The provider name
   * @return {string} Formatted symbol
   */
  static formatSymbol(symbol, provider) {
    if (!symbol) return symbol;
    
    const lowerProvider = (provider || "").toLowerCase().trim();
    
    switch (lowerProvider) {
      case "google":
        // Google Finance sometimes requires specific exchange symbols
        // For example, Korean stocks need KRX: prefix
        if (/^\d{6}$/.test(symbol) && !symbol.includes(':')) {
          return `KRX:${symbol}`;
        }
        return symbol;
        
      case "yahoo":
        // Yahoo Finance uses different symbols for Korean stocks (add .KS)
        if (/^\d{6}$/.test(symbol) && !symbol.includes('.')) {
          return `${symbol}.KS`;
        }
        return symbol;
        
      case "naver":
        // Naver Finance uses plain symbols
        return symbol;
        
      default:
        return symbol;
    }
  }
}

/**
 * Helper function to create or get the temp calculation sheet
 * @return {Sheet} The temporary sheet
 */
function getOrCreateTempSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEETS.TEMP);
  
  if (!sheet) {
    // Create the sheet if it doesn't exist
    sheet = ss.insertSheet(CONFIG.SHEETS.TEMP);
    
    // Add a header row
    sheet.getRange("A1:E1").setValues([["Formula", "Value", "Type", "Status", "Notes"]]);
    sheet.getRange("A1:E1").setFontWeight("bold");
    
    // Format the sheet
    sheet.setColumnWidth(1, 300); // Formula column wider
    sheet.setColumnWidth(2, 150); // Value column
    sheet.setColumnWidth(3, 100); // Type column
    sheet.setColumnWidth(4, 100); // Status column
    sheet.setColumnWidth(5, 300); // Notes column
    
    // Add a diagnostic section
    sheet.getRange("G1:H1").setValues([["Diagnostics", "Value"]]);
    sheet.getRange("G1:H1").setFontWeight("bold");
    
    // Hide the sheet by default (will be shown in diagnostic mode)
    sheet.hideSheet();
  }
  
  return sheet;
}

// Function to enable diagnostic mode for Google Finance data
function enableGoogleFinanceDiagnostics(enable = true) {
  const factory = new DataProviderFactory();
  const provider = factory.getProvider("google");
  
  if (provider && typeof provider.setDiagnosticMode === 'function') {
    provider.setDiagnosticMode(enable);
    return true;
  }
  
  return false;
} 