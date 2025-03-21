/**
 * Performance Dashboard Ticker - Price Service Module
 * 
 * This module provides functionality for fetching and managing price data.
 */

/**
 * Get prices for a ticker for all time periods
 * @param {string} symbol - The ticker symbol
 * @param {string} source - The data source
 * @param {DateCalculator} dateCalculator - The date calculator
 * @return {Object} Object containing all prices
 */
function getPrices(symbol, source, dateCalculator) {
  try {
    Logger.log(`${symbol} 가격 정보 조회 중 (${source})...`);
    
    // Enable diagnostic mode temporarily for debugging without showing alert
    const originalDiagnosticMode = isDiagnosticModeEnabled();
    enableDiagnosticMode(true, false); // Don't show alert when temporarily enabling
    
    // Initialize data provider
    const marketTimeManager = new MarketTimeManager();
    const dataProviderFactory = new DataProviderFactory(marketTimeManager);
    const dataProvider = dataProviderFactory.getProvider(source);
    
    if (!dataProvider) {
      throw new Error(`데이터 공급자를 찾을 수 없습니다: ${source}`);
    }
    
    // Get current price
    Logger.log(`${symbol}: 현재 가격 조회 중...`);
    const current = dataProvider.getPrice(symbol, dateCalculator.getReferenceDate());
    Logger.log(`${symbol}: 현재 가격 = ${current}`);
    
    if (typeof current !== 'number' || isNaN(current)) {
      throw new Error(`현재 가격을 가져올 수 없습니다: ${symbol} (${source})`);
    }
    
    // Get market region
    const marketRegion = marketTimeManager.getMarketRegion(source, symbol);
    Logger.log(`${symbol}: Market region = ${marketRegion}`);
    
    // Add annotations for the date
    const currentDate = dateCalculator.getReferenceDate();
    const mostRecentTradingDay = marketTimeManager.getMostRecentTradingDay(marketRegion, currentDate);
    
    // Get historical price data
    const weeklyDate = dateCalculator.getDateWeeksAgo(1);
    const monthlyDate = dateCalculator.getDateMonthsAgo(1);
    const ytdDate = dateCalculator.getYearStartDate();
    
    Logger.log(`${symbol}: 주간 날짜 = ${weeklyDate.toISOString()}`);
    Logger.log(`${symbol}: 월간 날짜 = ${monthlyDate.toISOString()}`);
    Logger.log(`${symbol}: 연초 날짜 = ${ytdDate.toISOString()}`);
    
    // Get prices for different time periods
    Logger.log(`${symbol}: 주간 가격 조회 중...`);
    let weekly = null;
    try {
      weekly = dataProvider.getHistoricalPrice(symbol, weeklyDate);
      Logger.log(`${symbol}: 주간 가격 = ${weekly}`);
    } catch (error) {
      Logger.log(`${symbol}: 주간 가격 조회 오류 - ${error.message}`);
      weekly = null;
    }
    
    Logger.log(`${symbol}: 월간 가격 조회 중...`);
    let monthly = null;
    try {
      monthly = dataProvider.getHistoricalPrice(symbol, monthlyDate);
      Logger.log(`${symbol}: 월간 가격 = ${monthly}`);
    } catch (error) {
      Logger.log(`${symbol}: 월간 가격 조회 오류 - ${error.message}`);
      monthly = null;
    }
    
    Logger.log(`${symbol}: 연초 가격 조회 중...`);
    let ytd = null;
    try {
      ytd = dataProvider.getHistoricalPrice(symbol, ytdDate);
      Logger.log(`${symbol}: 연초 가격 = ${ytd}`);
    } catch (error) {
      Logger.log(`${symbol}: 연초 가격 조회 오류 - ${error.message}`);
      ytd = null;
    }
    
    // Get highest price in 52 weeks
    Logger.log(`${symbol}: 52주 최고가 조회 중...`);
    let high = null;
    try {
      high = dataProvider.getHighPrice(symbol);
      Logger.log(`${symbol}: 52주 최고가 = ${high}`);
    } catch (error) {
      Logger.log(`${symbol}: 52주 최고가 조회 오류 - ${error.message}`);
      high = null;
    }
    
    // Debug the price data structure
    const prices = {
      current,
      high,
      weekly,
      monthly,
      ytd,
      referenceDate: mostRecentTradingDay
    };
    
    Logger.log(`${symbol} 가격 데이터 구조:`);
    Logger.log(JSON.stringify(prices, null, 2));
    
    // Restore original diagnostic mode without showing alert
    enableDiagnosticMode(originalDiagnosticMode, false);
    
    return prices;
  } catch (error) {
    // Ensure diagnostic mode is reset even if an error occurs
    try {
      enableDiagnosticMode(originalDiagnosticMode, false);
    } catch (e) {
      // Ignore errors in cleanup
    }
    
    Logger.log(`${symbol} 가격 정보 조회 실패: ${error.message}`);
    throw error;
  }
}

/**
 * Calculate returns for different time periods
 * @param {Object} prices - Object containing all prices
 * @return {Object} Object containing all returns
 */
function calculateReturns(prices) {
  if (!prices) {
    return { weekly: null, monthly: null, ytd: null, high: null };
  }
  
  return {
    weekly: calculateReturn(prices.current, prices.weekly),
    monthly: calculateReturn(prices.current, prices.monthly),
    ytd: calculateReturn(prices.current, prices.ytd),
    high: calculateReturn(prices.current, prices.high)
  };
}

/**
 * Calculate return as a percentage with sign
 * @param {number} current - Current price
 * @param {number} previous - Previous price
 * @return {number|null} Return percentage or null if invalid
 */
function calculateReturn(current, previous) {
  // Log detailed information for diagnostics
  if (isDiagnosticModeEnabled()) {
    Logger.log(`Calculating return - Current: ${current} (${typeof current}), Previous: ${previous} (${typeof previous})`);
  }
  
  // Check for invalid inputs
  if (current === null || previous === null || 
      current === undefined || previous === undefined ||
      isNaN(current) || isNaN(previous)) {
    Logger.log(`Invalid return calculation inputs - Current: ${current}, Previous: ${previous}`);
    return null;
  }
  
  // Convert to numbers if they are strings
  if (typeof current === 'string') current = parseFloat(current);
  if (typeof previous === 'string') previous = parseFloat(previous);
  
  // Additional check after conversion
  if (isNaN(current) || isNaN(previous)) {
    Logger.log(`Invalid numbers after conversion - Current: ${current}, Previous: ${previous}`);
    return null;
  }
  
  // Check for zero division
  if (previous === 0) {
    Logger.log(`Zero division error in return calculation - Previous price is zero`);
    return null;
  }
  
  // Check if prices are identical (with small tolerance for floating point)
  if (Math.abs(current - previous) < 0.000001) {
    return 0;
  }
  
  // Calculate percentage change
  const returnValue = ((current - previous) / previous) * 100;
  
  // Validate the result
  if (isNaN(returnValue) || !isFinite(returnValue)) {
    Logger.log(`Invalid calculation result: ${returnValue}`);
    return null;
  }
  
  // Log the result for diagnostics
  if (isDiagnosticModeEnabled()) {
    Logger.log(`Return calculation result: ${returnValue.toFixed(2)}%`);
  }
  
  return returnValue;
}

/**
 * Format a return value as a percentage string with sign
 * @param {number} returnValue - The return value to format
 * @return {string} Formatted return string
 */
function formatReturn(returnValue) {
  if (returnValue === null || returnValue === undefined || isNaN(returnValue)) {
    return "N/A";
  }
  
  const sign = returnValue > 0 ? "+" : "";
  return `${sign}${returnValue.toFixed(2)}%`;
}

/**
 * Format a price value as a string
 * @param {number} price - The price value to format
 * @return {string} Formatted price string
 */
function formatPrice(price) {
  if (price === null || price === undefined || isNaN(price)) {
    return "N/A";
  }
  
  return price.toFixed(2);
} 