/**
 * Performance Dashboard Ticker - Main Module
 * 
 * Contains the main application logic and entry points for the dashboard.
 */

// Main functionality
// Note: CONFIG, SS, DASHBOARD_SHEET, and REF_DATE_CELL are defined in Config.js
// Note: All data provider classes are defined in DataProviders.js
// Note: DateCalculator class is defined in DateUtils.js

/**
 * Global variable to track progress for the UI
 */
let PROGRESS_DATA = {
  percent: 0,
  message: "",
  type: "progress",
  tickers: []
};

/**
 * Lock utility for preventing race conditions
 */
function getScriptLock() {
  try {
    return LockService.getScriptLock();
  } catch (error) {
    Logger.log(`스크립트 잠금 생성 실패: ${error.message}`);
    return null;
  }
}

/**
 * Check if an update is already in progress
 * @return {boolean} True if update is in progress
 */
function isUpdateInProgress() {
  try {
    const props = PropertiesService.getScriptProperties();
    const updateStatus = props.getProperty('UPDATE_IN_PROGRESS');
    const updateStartTime = props.getProperty('UPDATE_START_TIME');
    
    // If no status is set, no update is in progress
    if (!updateStatus) return false;
    
    // Check if the update has been running for more than 10 minutes
    // This prevents a "stuck" lock if a script crashed
    if (updateStartTime) {
      const startTime = new Date(updateStartTime);
      const currentTime = new Date();
      const elapsedMinutes = (currentTime - startTime) / (1000 * 60);
      
      // If the update has been running for more than 10 minutes, consider it stuck
      // and allow a new update to start
      if (elapsedMinutes > 10) {
        Logger.log(`이전 업데이트가 ${elapsedMinutes.toFixed(1)}분 동안 실행 중이었습니다. 잠금을 해제합니다.`);
        clearUpdateStatus();
        return false;
      }
    }
    
    return updateStatus === 'true';
  } catch (error) {
    Logger.log(`업데이트 상태 확인 실패: ${error.message}`);
    return false;
  }
}

/**
 * Set update in progress status
 */
function setUpdateInProgress() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('UPDATE_IN_PROGRESS', 'true');
    props.setProperty('UPDATE_START_TIME', new Date().toString());
  } catch (error) {
    Logger.log(`업데이트 상태 설정 실패: ${error.message}`);
  }
}

/**
 * Clear update in progress status
 */
function clearUpdateStatus() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('UPDATE_IN_PROGRESS');
    props.deleteProperty('UPDATE_START_TIME');
  } catch (error) {
    Logger.log(`업데이트 상태 초기화 실패: ${error.message}`);
  }
}

/**
 * Display an error message in the spreadsheet
 * @param {string} errorMessage - The error message to display
 * @param {string} cell - Optional cell reference to display the error (defaults to DATA_RANGE_START)
 * @param {boolean} showAlert - Whether to show an alert dialog (defaults to true)
 * @param {Sheet} dashboardSheet - Optional dashboard sheet reference
 */
function displayError(errorMessage, cell, showAlert = true, dashboardSheet) {
  Logger.log(`오류: ${errorMessage}`);
  
  // Get dashboard sheet if not provided
  if (!dashboardSheet) {
    try {
      dashboardSheet = getDashboardSheet();
    } catch (e) {
      // If we can't get the dashboard sheet, just log the error and return
      Logger.log(`대시보드 시트를 찾을 수 없습니다: ${e.message}`);
      return;
    }
  }
  
  // Display error in sheet
  const targetCell = cell || CONFIG.DISPLAY.DATA_RANGE_START;
  dashboardSheet.getRange(targetCell).setValue(`⚠️ 오류: ${errorMessage}`);
  
  // Format the error cell to make it stand out
  dashboardSheet.getRange(targetCell).setBackground("#f4cccc");
  dashboardSheet.getRange(targetCell).setFontWeight("bold");
  
  // Also show alert if requested
  if (showAlert) {
    try {
      SpreadsheetApp.getUi().alert("오류", errorMessage, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      // If running in background or without UI, this might fail
      Logger.log("UI 알림을 표시할 수 없습니다: " + e.message);
    }
  }
}

/**
 * Updates the performance dashboard with the latest data
 * @param {boolean} adjustForBusinessDay - Whether to adjust for Korean business days
 */
function updatePerformanceDashboard(adjustForBusinessDay = true) {
  // Check if an update is already in progress
  if (isUpdateInProgress()) {
    const message = "다른 사용자가 이미 대시보드를 업데이트하고 있습니다. 잠시 후 다시 시도해 주세요.";
    Logger.log(message);
    SpreadsheetApp.getUi().alert(message);
    return;
  }
  
  // Try to get a lock for this operation
  const lock = getScriptLock();
  let lockAcquired = false;
  
  try {
    // Try to acquire a lock with timeout
    if (lock) {
      lockAcquired = lock.tryLock(10000); // wait up to 10 seconds
    }
    
    // If we couldn't get a lock, alert the user and exit
    if (!lockAcquired) {
      const message = "다른 사용자가 이미 대시보드를 업데이트하고 있습니다. 잠시 후 다시 시도해 주세요.";
      Logger.log(message);
      SpreadsheetApp.getUi().alert(message);
      return;
    }
    
    // Set update in progress flag
    setUpdateInProgress();
    
    // Start the actual update
    Logger.log("대시보드 업데이트를 시작합니다...");
    
    // Initialize if needed
    initialize();
    
    // Get the dashboard sheet
    let dashboardSheet;
    try {
      dashboardSheet = getDashboardSheet();
      if (!dashboardSheet) {
        throw new Error("대시보드 시트를 찾을 수 없습니다.");
      }
    } catch (error) {
      Logger.log(`대시보드 시트 가져오기 실패: ${error.message}`);
      displayError("대시보드 시트를 가져올 수 없습니다: " + error.message);
      return;
    }
    
    // Show loading indicator
    showLoadingIndicator("티커 데이터 로드 중...", dashboardSheet);
    
    // Get reference date
    let refDate;
    try {
      refDate = getReferenceDate(adjustForBusinessDay);
      if (!refDate) {
        throw new Error("참조일을 가져올 수 없습니다.");
      }
    } catch (error) {
      Logger.log(`참조일 가져오기 실패: ${error.message}`);
      hideLoadingIndicator(dashboardSheet);
      displayError("참조일을 가져올 수 없습니다: " + error.message, null, true, dashboardSheet);
      return;
    }
    
    // Load ticker data
    let tickers;
    try {
      tickers = getTickerData();
      if (!tickers || tickers.length === 0) {
        throw new Error("티커 데이터를 찾을 수 없습니다.");
      }
      Logger.log(`${tickers.length}개의 티커를 로드했습니다.`);
    } catch (error) {
      Logger.log(`티커 데이터 로드 실패: ${error.message}`);
      hideLoadingIndicator(dashboardSheet);
      displayError("티커 데이터를 로드할 수 없습니다: " + error.message, null, true, dashboardSheet);
      return;
    }
    
    // Update loading message
    updateLoadingIndicator("대시보드 준비 중...", dashboardSheet);
    
    // Clear dashboard and render header
    try {
      clearDashboard(dashboardSheet);
      renderHeader(dashboardSheet, refDate);
    } catch (error) {
      Logger.log(`대시보드 초기화 실패: ${error.message}`);
      hideLoadingIndicator(dashboardSheet);
      displayError("대시보드를 초기화할 수 없습니다: " + error.message, null, true, dashboardSheet);
      return;
    }
    
    // Process and render tickers
    try {
      processAndRenderTickersProgressively(tickers, refDate, dashboardSheet);
    } catch (error) {
      Logger.log(`티커 처리 실패: ${error.message}`);
      hideLoadingIndicator(dashboardSheet);
      displayError("티커를 처리할 수 없습니다: " + error.message, null, true, dashboardSheet);
      return;
    }
    
    // Format dashboard
    try {
      formatDashboard(dashboardSheet);
    } catch (error) {
      Logger.log(`대시보드 서식 지정 실패: ${error.message}`);
      // Continue despite formatting errors
    }
    
    // Hide loading indicator
    hideLoadingIndicator(dashboardSheet);
    
    Logger.log("대시보드 업데이트 완료");
  } catch (error) {
    Logger.log(`대시보드 업데이트 중 오류 발생: ${error.message}`);
    // Try to hide loading indicator if there was an error
    try {
      if (dashboardSheet) {
        hideLoadingIndicator(dashboardSheet);
      }
    } catch (e) {
      // Ignore errors in error handling
    }
    
    displayError("대시보드 업데이트 중 오류가 발생했습니다: " + error.message);
  } finally {
    // Clear update in progress flag
    clearUpdateStatus();
    
    // Release the lock if we acquired it
    if (lockAcquired && lock) {
      try {
        lock.releaseLock();
      } catch (e) {
        Logger.log(`잠금 해제 실패: ${e.message}`);
      }
    }
  }
}

/**
 * Update progress data
 * @param {number} percent - Progress percentage (0-100)
 * @param {string} message - Progress message
 */
function updateProgress(percent, message) {
  PROGRESS_DATA = {
    percent: percent,
    message: message,
    type: "progress",
    tickers: PROGRESS_DATA.tickers
  };
}

/**
 * Get current progress data
 * @return {Object} Current progress data
 */
function getProgress() {
  return PROGRESS_DATA;
}

/**
 * Look up a specific symbol for the sidebar
 * @param {string} symbol - Ticker symbol to look up
 * @param {string} source - Data source (google, yahoo, naver)
 * @return {Object} The price and return data
 */
function lookupSymbol(symbol, source) {
  try {
    if (!symbol) {
      return {error: "심볼이 제공되지 않았습니다."};
    }
    
    if (!source) {
      return {error: "데이터 소스가 지정되지 않았습니다."};
    }
    
    Logger.log(`단일 티커 조회: ${symbol} (${source})`);
    
  const referenceDate = getReferenceDate();
    const dateCalculator = new DateCalculator(referenceDate);
    
    const prices = getPrices(symbol, source, dateCalculator);
    const returns = calculateReturns(prices);
    
    return {
      symbol: symbol,
      current: prices.current,
      high: prices.high,
      weekly: returns.weekly,
      monthly: returns.monthly,
      ytd: returns.ytd,
      error: null
    };
  } catch (error) {
    Logger.log(`티커 조회 오류: ${error.message}`);
    return {
      symbol: symbol,
      error: `조회 중 오류가 발생했습니다: ${error.message}`
    };
  }
}

/**
 * Process and render ticker data progressively
 * @param {Array} tickerList - List of tickers to process
 * @param {Date} referenceDate - Reference date for calculations
 * @param {Sheet} dashboardSheet - The dashboard sheet
 */
function processAndRenderTickersProgressively(tickerList, referenceDate, dashboardSheet) {
  // If dashboard sheet isn't provided, try to get it directly
  if (!dashboardSheet) {
    dashboardSheet = getDashboardSheet();
  }
  
  const dateCalculator = new DateCalculator(referenceDate);
  const data = [];
  const errors = [];
  
  // Reset progress tracking
  PROGRESS_DATA = {
    percent: 0,
    message: "시작 중...",
    type: "progress",
    tickers: []
  };
  
  // First, prepare the sheet with empty rows for all tickers
  // This creates a visual placeholder for each ticker
  const initialData = tickerList.map(ticker => [
    ticker.name,
    "로딩 중...",
    "로딩 중...",
    "로딩 중...",
    "로딩 중..."
  ]);
  
  // Render the initial placeholder data
  renderData(initialData, dashboardSheet);
  
  // Get the number of tickers to track progress
  const totalTickers = tickerList.length;
  
  // Process each ticker and update the UI after each one
  for (let i = 0; i < tickerList.length; i++) {
    const ticker = tickerList[i];
    try {
    const { name, ticker: symbol, source } = ticker;
      
      // Calculate and update progress
      const progressPercent = Math.round((i / totalTickers) * 100);
      updateProgress(progressPercent, `처리 중 (${i+1}/${totalTickers}): ${name}`);
      
      // Update the loading message to show which ticker is being processed
      updateLoadingIndicator(`처리 중 (${i+1}/${totalTickers}): ${name}`, dashboardSheet);
      
      Logger.log(`처리 중: ${name} (${symbol}) - 소스: ${source}`);
      
      // Special handling for Google Finance (which may require different format)
      if (source.toLowerCase() === 'google') {
        Logger.log(`구글 파이낸스 데이터 처리 중 - 심볼: ${symbol}`);
      }
      
    const prices = getPrices(symbol, source, dateCalculator);
      
      // Log all prices for debugging
      logPrice("현재가", prices.current, symbol);
      logPrice("일주일 전 가격", prices.weekAgo, symbol);
      logPrice("한달 전 가격", prices.monthAgo, symbol);
      logPrice("연초 가격", prices.ytd, symbol);
      logPrice("52주 고점", prices.high, symbol);
      
    const returns = calculateReturns(prices);
    
      // Add to progress data
      PROGRESS_DATA.tickers.push({
        name: name,
        symbol: symbol,
        status: "completed"
      });
      
      // Update just this ticker's row in the sheet
      updateTickerRow(i, [name, returns.weekly, returns.monthly, returns.ytd, returns.high], dashboardSheet);
      
      // Also store in our data array for later reference
    data.push([name, returns.weekly, returns.monthly, returns.ytd, returns.high]);
      
      Logger.log(`처리 완료: ${name}: ${JSON.stringify(returns)}`);
    } catch (error) {
      // Record error but continue processing other tickers
      Logger.log(`티커 처리 오류 - ${ticker.name}: ${error.message}`);
      errors.push(`${ticker.name || ticker.ticker}: ${error.message}`);
      
      // Add to progress data
      PROGRESS_DATA.tickers.push({
        name: ticker.name,
        symbol: ticker.ticker,
        status: "error",
        error: error.message
      });
      
      // Update the row with error indicators
      updateTickerRow(i, [
        ticker.name, 
        CONFIG.STATUS.ERROR, 
        CONFIG.STATUS.ERROR, 
        CONFIG.STATUS.ERROR, 
        CONFIG.STATUS.ERROR
      ], dashboardSheet);
      
      // Also add to our data array
      data.push([
        ticker.name, 
        CONFIG.STATUS.ERROR, 
        CONFIG.STATUS.ERROR, 
        CONFIG.STATUS.ERROR, 
        CONFIG.STATUS.ERROR
      ]);
    }
    
    // Small delay to allow UI to update
    Utilities.sleep(50);
  }
  
  // Update progress to 100%
  updateProgress(100, "완료");
  
  // Apply final formatting to the entire dataset at once
  applyReturnFormatting(data, dashboardSheet);
  
  // Display errors in a separate cell if there were any issues
  if (errors.length > 0) {
    const errorCell = "A" + (data.length + 5);
    const errorMessage = `일부 티커에서 문제가 발생했습니다:\n${errors.join("\n")}`;
    dashboardSheet.getRange(errorCell).setValue(errorMessage);
    dashboardSheet.getRange(errorCell).setBackground("#fff2cc");
  }
  
  // Add timestamp
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  dashboardSheet.getRange("A" + (data.length + 4)).setValue(`마지막 업데이트: ${timestamp}`);
  
  // Clear loading indicator
  hideLoadingIndicator(dashboardSheet);
}

/**
 * Update a single ticker row in the sheet
 * @param {number} rowIndex - Index of the row in the data
 * @param {Array} rowData - The data for this row
 * @param {Sheet} dashboardSheet - The dashboard sheet
 */
function updateTickerRow(rowIndex, rowData, dashboardSheet) {
  const targetRow = rowIndex + 3; // +3 to account for headers
  
  // Check if the rowData has loading indicators
  const isLoading = rowData[1] === "로딩 중...";
  
  // Write data to the sheet
  dashboardSheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
  
  // If this is a loading row, style it accordingly
  if (isLoading) {
    // Gray out the loading text and add a subtle loading background
    dashboardSheet.getRange(targetRow, 2, 1, 4).setFontColor("#999999");
    dashboardSheet.getRange(targetRow, 1, 1, 5).setBackground("#f5f5f5");
    // Make the name bold regardless
    dashboardSheet.getRange(targetRow, 1).setFontWeight("bold");
    return;
  }
  
  // Center align all the return cells
  dashboardSheet.getRange(targetRow, 2, 1, 4).setHorizontalAlignment("center");
  
  // Left align the ticker name cell
  dashboardSheet.getRange(targetRow, 1).setHorizontalAlignment("left");
  
  // Apply immediate formatting based on return values
  for (let i = 1; i < rowData.length; i++) { // For each column (skip name column)
    const column = i + 1; // Adjust for 1-indexed columns
    const cell = dashboardSheet.getRange(targetRow, column);
    const value = rowData[i];
    
    // Convert to string to ensure startsWith will work
    const valueStr = String(value);
    
    // Clear any previous background and formatting
    cell.setBackground(null);
    cell.setFontWeight("normal");
    
    // Apply formatting based on the value
    if (value === CONFIG.STATUS.ERROR || 
        value === CONFIG.STATUS.NO_DATA || 
        value === CONFIG.STATUS.CALC_ERROR) {
      // Error or no data - light gray
      cell.setFontColor("#999999");
      cell.setFontStyle("italic");
    } else if (valueStr.startsWith("+")) {
      // Positive return - green with subtle background
      cell.setFontColor("#006100");
      cell.setBackground("#e6f4ea"); // Light green background
      
      // If it's a significant positive return (>5%), make it bold
      if (parseFloat(valueStr) > 5) {
        cell.setFontWeight("bold");
      }
    } else if (valueStr.startsWith("-")) {
      // Negative return - red with subtle background
      cell.setFontColor("#990000");
      cell.setBackground("#fce8e6"); // Light red background
      
      // If it's a significant negative return (<-5%), make it bold
      if (parseFloat(valueStr) < -5) {
        cell.setFontWeight("bold");
      }
    } else {
      // Neutral or unknown - black
      cell.setFontColor("#000000");
    }
  }
  
  // Style the name column
  dashboardSheet.getRange(targetRow, 1).setFontWeight("bold");
  dashboardSheet.getRange(targetRow, 1).setBackground(null);
}

/**
 * Show a loading indicator in the spreadsheet (simplified version without merged cells)
 * @param {string} message - The loading message to display
 * @param {Sheet} dashboardSheet - The dashboard sheet
 */
function showLoadingIndicator(message, dashboardSheet) {
  // If dashboard sheet isn't provided, try to get it directly
  if (!dashboardSheet) {
    dashboardSheet = getDashboardSheet();
  }
  
  // Create a loading cell at the top of the sheet (without merging)
  const loadingCell = dashboardSheet.getRange("A1");
  loadingCell.setValue(`⏳ ${message}`);
  loadingCell.setBackground("#e6f2ff");
  loadingCell.setFontWeight("bold");
  
  // Force the spreadsheet to update
  SpreadsheetApp.flush();
}

/**
 * Update the loading message
 * @param {string} message - New loading message
 * @param {Sheet} dashboardSheet - The dashboard sheet
 */
function updateLoadingIndicator(message, dashboardSheet) {
  // If dashboard sheet isn't provided, try to get it directly
  if (!dashboardSheet) {
    dashboardSheet = getDashboardSheet();
  }
  
  dashboardSheet.getRange("A1").setValue(`⏳ ${message}`);
  
  // Force the spreadsheet to update
  SpreadsheetApp.flush();
}

/**
 * Hide the loading indicator
 * @param {Sheet} dashboardSheet - The dashboard sheet
 */
function hideLoadingIndicator(dashboardSheet) {
  // If dashboard sheet isn't provided, try to get it directly
  if (!dashboardSheet) {
    dashboardSheet = getDashboardSheet();
  }
  
  try {
    // Simply clear and reset the cell styling
    dashboardSheet.getRange("A1").setValue(""); // Clear the loading indicator
    dashboardSheet.getRange("A1").setBackground(null);
    dashboardSheet.getRange("A1").setFontWeight("normal");
    
    // Force the spreadsheet to update
    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log("로딩 표시기를 제거하는 중 오류가 발생했습니다: " + error.message);
    // Try minimal cleanup even if error occurred
    try {
      dashboardSheet.getRange("A1").setValue("");
    } catch (e) {
      // Ignore secondary errors
    }
  }
}

/**
 * Helper function to log prices with appropriate formatting
 * @param {string} description - Description of the price
 * @param {number|string} price - The price value
 * @param {string} symbol - The ticker symbol
 */
function logPrice(description, price, symbol) {
  if (isNaN(price)) {
    Logger.log(`${description} for ${symbol}: NaN (invalid number)`);
  } else if (price === 0) {
    Logger.log(`${description} for ${symbol}: 0 (possible data retrieval issue)`);
  } else if (price === "데이터 없음") {
    Logger.log(`${description} for ${symbol}: 데이터 없음 (no data available)`);
  } else {
    Logger.log(`${description} for ${symbol}: ${price}`);
  }
}

/**
 * Calculate returns for prices
 * @param {Object} prices - Object with current, weekAgo, monthAgo, ytd, and high prices
 * @return {Object} Returns for each timeframe
 */
function calculateReturns(prices) {
  return {
    weekly: calculateReturn(prices.current, prices.weekAgo),
    monthly: calculateReturn(prices.current, prices.monthAgo),
    ytd: calculateReturn(prices.current, prices.ytd),
    high: calculateReturn(prices.current, prices.high)
  };
}

/**
 * Calculate a return percentage
 * @param {number} current - Current price
 * @param {number} past - Past price
 * @return {string} Formatted return percentage
 */
function calculateReturn(current, past) {
  try {
    // Add detailed logging for debugging
    Logger.log(`계산 시작 - 현재가: ${current} (${typeof current}), 과거가: ${past} (${typeof past})`);
    
    // Check for invalid inputs
    if (current === CONFIG.STATUS.NO_DATA || past === CONFIG.STATUS.NO_DATA) {
      Logger.log(`NO_DATA 발견 - 현재가: ${current}, 과거가: ${past}`);
      return CONFIG.STATUS.NO_DATA;
    }
    
    // Convert to numbers if they are strings (and not error indicators)
    if (typeof current === 'string' && !isNaN(current)) {
      current = parseFloat(current);
      Logger.log(`현재가 문자열에서 숫자로 변환: ${current}`);
    }
    
    if (typeof past === 'string' && !isNaN(past)) {
      past = parseFloat(past);
      Logger.log(`과거가 문자열에서 숫자로 변환: ${past}`);
    }
    
    // Check for invalid numbers
    if (isNaN(current)) {
      Logger.log(`유효하지 않은 현재가: ${current}`);
      return CONFIG.STATUS.CALC_ERROR;
    }
    
    if (isNaN(past)) {
      Logger.log(`유효하지 않은 과거가: ${past}`);
      return CONFIG.STATUS.CALC_ERROR;
    }
    
    if (past === 0) {
      Logger.log(`과거가가 0입니다 - 나눗셈 오류 방지`);
      return CONFIG.STATUS.CALC_ERROR;
    }
    
    // Special case: if current and past are the same (happens with future dates when all prices are current)
    if (current === past) {
      Logger.log(`현재가와 과거가가 동일합니다 (${current}) - 미래 날짜 또는 데이터 문제로 인해 발생 가능`);
      return "0.00%";
    }
    
    // Calculate the return
    const returnValue = (current / past - 1) * 100;
    Logger.log(`수익률 계산: (${current} / ${past} - 1) * 100 = ${returnValue}%`);
    
    // Format the return with proper sign and rounding
    if (isNaN(returnValue)) {
      Logger.log(`NaN 결과: 현재=${current}, 과거=${past}`);
      return CONFIG.STATUS.CALC_ERROR;
    }
    
    // Format with + sign for positive returns
    const sign = returnValue > 0 ? '+' : '';
    const formattedReturn = `${sign}${returnValue.toFixed(2)}%`;
    Logger.log(`최종 서식화된 수익률: ${formattedReturn}`);
    
    return formattedReturn;
  } catch (error) {
    Logger.log(`계산 오류: ${error.message}, 스택: ${error.stack}`);
    return CONFIG.STATUS.CALC_ERROR;
  }
}

/**
 * Render data to the dashboard
 * @param {Array} data - The data to render
 * @param {Sheet} dashboardSheet - The dashboard sheet
 */
function renderData(data, dashboardSheet) {
  if (data.length === 0) return;
  
  // Write the data starting from row 3
  dashboardSheet.getRange(3, 1, data.length, data[0].length).setValues(data);
  
  // Apply formatting
  applyReturnFormatting(data, dashboardSheet);
}

/**
 * Apply color formatting to the rendered data
 * @param {Array} data - The data that was rendered
 * @param {Sheet} dashboardSheet - The dashboard sheet
 */
function applyReturnFormatting(data, dashboardSheet) {
  if (data.length === 0) return;
  
  // For each column of returns (columns B, C, D, E in the displayed data)
  const columns = ['B', 'C', 'D', 'E'];
  const startRow = 3; // First row of data after header
  
  columns.forEach(column => {
    for (let i = 0; i < data.length; i++) {
      const row = startRow + i;
      const cell = dashboardSheet.getRange(`${column}${row}`);
      const value = cell.getValue();
      
      // Convert value to string to ensure startsWith will work
      const valueStr = String(value);
      
      // Apply formatting based on the value
      if (value === CONFIG.STATUS.ERROR || 
          value === CONFIG.STATUS.NO_DATA || 
          value === CONFIG.STATUS.CALC_ERROR) {
        // Error or no data - light gray
        cell.setFontColor("#999999");
      } else if (valueStr.startsWith("+")) {
        // Positive return - green
        cell.setFontColor("#006100");
      } else if (valueStr.startsWith("-")) {
        // Negative return - red
        cell.setFontColor("#990000");
      } else {
        // Neutral or unknown - black
        cell.setFontColor("#000000");
      }
    }
  });
  
  // Make the name column (column A) bold
  for (let i = 0; i < data.length; i++) {
    const row = startRow + i;
    dashboardSheet.getRange(`A${row}`).setFontWeight("bold");
  }
}

/**
 * Render the header row
 * @param {Sheet} dashboardSheet - The dashboard sheet
 * @param {Date} referenceDate - Reference date to display
 */
function renderHeader(dashboardSheet, referenceDate) {
  // Format reference date for display 
  const formattedDate = Utilities.formatDate(referenceDate, "Asia/Seoul", "yyyy년 MM월 dd일");
  const dayOfWeek = ["일", "월", "화", "수", "목", "금", "토"][referenceDate.getDay()];
  const dateDisplay = `${formattedDate} (${dayOfWeek})`;
  
  // Check if the date is a weekend or Korean holiday
  const isWeekend = (referenceDate.getDay() === 0 || referenceDate.getDay() === 6);
  let marketStatus = "";
  
  if (isWeekend) {
    marketStatus = " - 주말 (시장 휴장)";
  } else if (isKoreanHoliday(referenceDate)) {
    marketStatus = " - 공휴일 (시장 휴장)";
  }
  
  // Make sure the header matches the data columns exactly
  const headers = [
    "지수",             // Column A - Index name
    "주간 수익률",      // Column B - Weekly return 
    "월간 수익률",      // Column C - Monthly return
    "YTD 수익률",       // Column D - YTD return
    "고점 대비 수익률"  // Column E - Return vs High
  ];
  
  // Set header values in row 2 (A2:E2)
  dashboardSheet.getRange(2, 1, 1, headers.length).setValues([headers]);
  
  // Add reference date info in row 1
  const dateCell = dashboardSheet.getRange("A1");
  dateCell.setValue(`기준일자: ${dateDisplay}${marketStatus}`);
  
  // Style the date cell
  dateCell.setFontWeight("bold");
  dateCell.setFontColor("#1a3370");
  dateCell.setBackground("#eef4ff");
  
  // Format the header with a more noticeable style
  const headerRange = dashboardSheet.getRange(2, 1, 1, headers.length);
  
  // Set gradient background for header
  headerRange.setBackground("#e6eefa");
  headerRange.setFontWeight("bold");
  headerRange.setFontColor("#1a3370");
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  
  // Add borders to the header
  headerRange.setBorder(
    true, true, true, true, false, false,
    "#c0cde3", SpreadsheetApp.BorderStyle.SOLID
  );
  
  // Apply different alignment for the first column (left-aligned)
  dashboardSheet.getRange(2, 1).setHorizontalAlignment("left");
  
  // Add shading to differentiate between different return periods
  dashboardSheet.getRange(2, 2).setBackground("#d9e8fc"); // 주간 수익률
  dashboardSheet.getRange(2, 3).setBackground("#d9e2fc"); // 월간 수익률
  dashboardSheet.getRange(2, 4).setBackground("#d9dcfc"); // YTD 수익률
  dashboardSheet.getRange(2, 5).setBackground("#e0d9fc"); // 고점 대비 수익률
  
  // Add a bottom border to clearly separate header from data
  headerRange.setBorder(
    false, false, true, false, false, false,
    "#8897bd", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
  );
  
  // Set row height for the header row
  dashboardSheet.setRowHeight(2, 30);
}

/**
 * Test function for debugging
 */
function test() {
  const provider = new GoogleFinanceProvider();
  const result = provider.getPrice("INDEXSP:.INX", new Date("2023-03-18"));
  Logger.log(result);
}

/**
 * Test the data providers
 */
function testDataProviders() {
  Logger.log("========= Testing Data Providers =========");
  
  // Test Naver Data Provider
  Logger.log("--- Testing Naver Finance Provider ---");
  const naverProvider = new NaverFinanceProvider();
  
  // KOSPI index
  const kospiSymbol = "KOSPI";
  const kospiPrice = naverProvider.getPrice(kospiSymbol, new Date());
  Logger.log(`Naver KOSPI (${kospiSymbol}) price: ${kospiPrice}`);
  
  // KOSDAQ index
  const kosdaqSymbol = "KOSDAQ";
  const kosdaqPrice = naverProvider.getPrice(kosdaqSymbol, new Date());
  Logger.log(`Naver KOSDAQ (${kosdaqSymbol}) price: ${kosdaqPrice}`);
  
  const kosdaqHigh = naverProvider.getHighPrice(kosdaqSymbol);
  Logger.log(`Naver KOSDAQ (${kosdaqSymbol}) 52-week high: ${kosdaqHigh}`);
  
  // Samsung Electronics
  const samsungSymbol = "005930";
  const samsungPrice = naverProvider.getPrice(samsungSymbol, new Date());
  Logger.log(`Naver Samsung (${samsungSymbol}) price: ${samsungPrice}`);
  
  const samsungHigh = naverProvider.getHighPrice(samsungSymbol);
  Logger.log(`Naver Samsung (${samsungSymbol}) 52-week high: ${samsungHigh}`);
  
  // Test Yahoo Data Provider
  Logger.log("--- Testing Yahoo Finance Provider ---");
  const yahooProvider = new YahooFinanceProvider();
  
  // Apple Inc.
  const appleSymbol = "AAPL";
  const applePrice = yahooProvider.getPrice(appleSymbol, new Date());
  Logger.log(`Yahoo Apple (${appleSymbol}) current price: ${applePrice}`);
  
  // Test historical price from a month ago
  const monthAgo = new Date();
  monthAgo.setMonth(monthAgo.getMonth() - 1);
  const appleHistorical = yahooProvider.getPrice(appleSymbol, monthAgo);
  Logger.log(`Yahoo Apple (${appleSymbol}) price on ${monthAgo.toISOString().split('T')[0]}: ${appleHistorical}`);
  
  const appleHigh = yahooProvider.getHighPrice(appleSymbol);
  Logger.log(`Yahoo Apple (${appleSymbol}) 52-week high: ${appleHigh}`);
  
  // S&P 500 Index
  const spSymbol = "^GSPC";
  const spPrice = yahooProvider.getPrice(spSymbol, new Date());
  Logger.log(`Yahoo S&P 500 (${spSymbol}) current price: ${spPrice}`);
  
  const spHigh = yahooProvider.getHighPrice(spSymbol);
  Logger.log(`Yahoo S&P 500 (${spSymbol}) 52-week high: ${spHigh}`);
  
  Logger.log("========= Testing Completed =========");
}

/**
 * Create and show a sidebar with interactive controls
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('성과 대시보드 컨트롤')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Create the menu when the spreadsheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 성과 대시보드')
    .addItem('대시보드 업데이트', 'updatePerformanceDashboard')
    .addSeparator()
    .addItem('참조일 설정', 'showSidebar')
    .addItem('티커 관리', 'openTickerSheet')
    .addSeparator()
    .addItem('진단 실행', 'runDiagnostics')
    .addItem('도움말', 'showHelp')
    .addToUi();
}

/**
 * Display diagnostic information to help users troubleshoot problems
 */
function runDiagnostics() {
  try {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check for tickers sheet
    const tickerSheet = sheet.getSheetByName("tickers");
    let tickerStatus = "찾을 수 없음";
    let tickerCount = 0;
    
    if (tickerSheet) {
      const data = tickerSheet.getDataRange().getValues();
      tickerStatus = "확인됨";
      tickerCount = data.length > 1 ? data.length - 1 : 0; // Subtract header row
    }
    
    // Check for reference date
    const refDate = getRefDateValue();
    let dateStatus = "찾을 수 없음";
    
    if (refDate) {
      if (Object.prototype.toString.call(refDate) === '[object Date]' && !isNaN(refDate.getTime())) {
        dateStatus = Utilities.formatDate(refDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        dateStatus = "유효하지 않은 날짜 형식";
      }
    }
    
    // Run sample provider tests
    let naverStatus = "테스트 중...";
    let yahooStatus = "테스트 중...";
    
    try {
      const naverProvider = new NaverFinanceProvider();
      const kospiPrice = naverProvider.getPrice("KOSPI", new Date());
      naverStatus = kospiPrice > 0 ? "정상" : "오류 (데이터 없음)";
    } catch (e) {
      naverStatus = `오류 (${e.message})`;
    }
    
    try {
      const yahooProvider = new YahooFinanceProvider();
      const spPrice = yahooProvider.getPrice("^GSPC", new Date());
      yahooStatus = spPrice > 0 ? "정상" : "오류 (데이터 없음)";
    } catch (e) {
      yahooStatus = `오류 (${e.message})`;
    }
    
    // Display results
    const message = 
      `진단 결과:\n\n` +
      `- 티커 시트: ${tickerStatus} (항목 수: ${tickerCount})\n` +
      `- 참조 날짜: ${dateStatus}\n` +
      `- 네이버 데이터 제공자: ${naverStatus}\n` +
      `- 야후 데이터 제공자: ${yahooStatus}\n\n` +
      `오류가 발생한 경우 아래 단계를 확인하세요:\n` +
      `1. A1 셀에 유효한 날짜가 있는지 확인하세요.\n` +
      `2. 티커 시트에 올바른 데이터가 있는지 확인하세요.\n` +
      `3. 각 티커 항목에 이름, 심볼, 데이터소스(google, yahoo, naver)가 포함되어 있는지 확인하세요.`;
    
    ui.alert("진단 결과", message, ui.ButtonSet.OK);
  } catch (error) {
    displayError(`진단 중 오류가 발생했습니다: ${error.message}`);
  }
}

/**
 * Show the help dialog
 */
function showHelp() {
  const htmlOutput = HtmlService
    .createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h2>성과 대시보드 도움말</h2>
        
        <h3>대시보드 업데이트</h3>
        <p>'업데이트' 메뉴를 클릭하여 모든 티커의 최신 가격 정보로 대시보드를 업데이트합니다. 이 작업은 티커의 수에 따라 시간이 소요될 수 있습니다.</p>
        
        <h3>티커 관리</h3>
        <p>'티커' 시트에서 추적하려는 종목을 관리할 수 있습니다:</p>
        <ul>
          <li><strong>이름</strong>: 종목의 이름 (예: "삼성전자")</li>
          <li><strong>심볼</strong>: 종목의 티커 심볼 (예: "005930" 또는 "AAPL")</li>
          <li><strong>데이터소스</strong>: 데이터를 가져올 소스 (google, yahoo, naver)</li>
        </ul>
        
        <h3>참조일 관리</h3>
        <p>대시보드는 기본적으로 가장 최근의 영업일을 참조일로 사용합니다. 직접 참조일을 설정하려면 'MM/DD/YYYY' 형식으로 입력하세요.</p>
        
        <h3>지원되는 데이터 소스</h3>
        <ul>
          <li><strong>Google Finance</strong>: 대부분의 글로벌 주식과 지수</li>
          <li><strong>Yahoo Finance</strong>: 대부분의 글로벌 주식과 ETF</li>
          <li><strong>Naver Finance</strong>: 한국 종목과 KOSPI/KOSDAQ 지수</li>
        </ul>
        
        <h3>참고 사항</h3>
        <ul>
          <li>대시보드 업데이트는 한 번에 한 명의 사용자만 실행할 수 있습니다.</li>
          <li>주말과 공휴일에는 최근 영업일의 데이터가 사용됩니다.</li>
          <li>대시보드에는 최대 200개의 티커를 추가할 수 있습니다.</li>
          <li>모든 수익률은 현지 통화 기준으로 계산됩니다.</li>
        </ul>
      </div>
    `)
    .setWidth(600)
    .setHeight(500);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, '성과 대시보드 도움말');
}

/**
 * Opens the tickers sheet for editing
 */
function openTickerSheet() {
  const sheet = getOrCreateTickerSheet();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  SpreadsheetApp.getUi().alert('티커 시트에서 추적하려는 종목을 추가하거나 수정하세요.');
}

/**
 * Clear the dashboard and reset formatting
 * @param {Sheet} dashboardSheet - The dashboard sheet
 */
function clearDashboard(dashboardSheet) {
  try {
    // Handle case where dashboardSheet is not provided
    if (!dashboardSheet) {
      try {
        dashboardSheet = getDashboardSheet();
      } catch (e) {
        Logger.log(`대시보드 시트를 가져오는 중 오류가 발생했습니다: ${e.message}`);
        return;
      }
    }
    
    // Get the last row that might have data
    const lastRow = dashboardSheet.getLastRow();
    
    // Clear the data area (starting from row 3)
    if (lastRow > 2) {
      dashboardSheet.getRange(3, 1, lastRow - 2, 10).clearContent();
      dashboardSheet.getRange(3, 1, lastRow - 2, 10).setBackground(null);
      dashboardSheet.getRange(3, 1, lastRow - 2, 10).setFontWeight("normal");
    }
    
    // Reset any special formatting
    try {
      // Use safe access to CONFIG
      const dataRangeStart = CONFIG && CONFIG.DISPLAY && CONFIG.DISPLAY.DATA_RANGE_START 
        ? CONFIG.DISPLAY.DATA_RANGE_START 
        : "A3";
      
      dashboardSheet.getRange(dataRangeStart).setBackground(null);
      dashboardSheet.getRange(dataRangeStart).setFontWeight("normal");
    } catch (e) {
      Logger.log(`특수 서식을 재설정하는 중 오류가 발생했습니다: ${e.message}`);
      // Continue with the function even if this part fails
    }
  } catch (error) {
    Logger.log(`대시보드 지우기 중 오류가 발생했습니다: ${error.message}`);
    // Let the caller handle this error
    throw error;
  }
}

/**
 * Gets the reference date from the spreadsheet
 * @param {boolean} adjustForBusinessDay - Whether to adjust for Korean business days
 * @return {Date} The reference date
 */
function getReferenceDate(adjustForBusinessDay = false) {
  // Get the reference date from the spreadsheet
  let refDate = getRefDateValue();
  
  // Check if the date is empty or invalid
  if (!refDate) {
    // Use today's date as the default
    Logger.log("참조일이 비어 있거나 유효하지 않습니다. 오늘 날짜를 사용합니다.");
    refDate = new Date();
  }
  
  // Convert to date object if it's not already
  if (!(refDate instanceof Date)) {
    try {
      refDate = new Date(refDate);
    } catch (e) {
      Logger.log("참조일 형식이 유효하지 않습니다. 오늘 날짜를 사용합니다.");
      refDate = new Date();
    }
  }
  
  // Check if the date is valid
  if (isNaN(refDate.getTime())) {
    Logger.log("참조일이 유효하지 않습니다. 오늘 날짜를 사용합니다.");
    refDate = new Date();
  }
  
  // Adjust for Korean business days if requested
  if (adjustForBusinessDay) {
    const lastBusinessDay = getLastKoreanBusinessDay();
    if (!datesEqual(lastBusinessDay, refDate)) {
      Logger.log(`참조일 조정: ${formatDate(refDate)} → ${formatDate(lastBusinessDay)} (마지막 영업일 기준)`);
      refDate = lastBusinessDay;
    }
  }
  
  return refDate;
}

/**
 * Get the current reference date set in the sidebar
 */
function getCurrentReferenceDate() {
  try {
    const dateValue = getRefDateValue();
    if (!dateValue) {
      return "";
    }
    
    const date = new Date(dateValue);
    // Format as YYYY-MM-DD
    return Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  } catch (error) {
    Logger.log("참조 날짜를 가져오는 중 오류가 발생했습니다: " + error.message);
    return "";
  }
}

/**
 * Set the reference date from the sidebar
 * @param {string} dateStr - Date string in format YYYY-MM-DD
 */
function setReferenceDate(dateStr) {
  try {
    // Parse the date string
    const date = new Date(dateStr);
    
    // Validate the date
    if (isNaN(date.getTime())) {
      const errorMsg = "유효하지 않은 날짜 형식입니다. YYYY-MM-DD 형식을 사용하세요.";
      SpreadsheetApp.getUi().alert("오류", errorMsg, SpreadsheetApp.getUi().ButtonSet.OK);
      return {"success": false, "message": errorMsg};
    }
    
    // Set the reference date
    const success = setRefDateValue(date);
    if (!success) {
      const errorMsg = "참조 날짜를 설정하는 중 오류가 발생했습니다.";
      return {"success": false, "message": errorMsg};
    }
    
    // Format the date for confirmation
    const formattedDate = Utilities.formatDate(date, "GMT+9", "yyyy년 MM월 dd일");
    Logger.log("참조 날짜가 설정되었습니다: " + formattedDate);
    
    return {
      "success": true, 
      "message": "참조 날짜가 설정되었습니다: " + formattedDate,
      "date": Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd")
    };
  } catch (error) {
    Logger.log("참조 날짜를 설정하는 중 오류가 발생했습니다: " + error.message);
    return {"success": false, "message": "참조 날짜를 설정하는 중 오류가 발생했습니다: " + error.message};
  }
}

/**
 * Check if the tickers sheet exists and create it if necessary
 * @return {Sheet} The ticker sheet
 */
function getOrCreateTickerSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let tickerSheet = ss.getSheetByName(CONFIG.SHEETS.TICKERS);
    
    if (!tickerSheet) {
      tickerSheet = ss.insertSheet(CONFIG.SHEETS.TICKERS);
      
      // Set up header row
      const headers = ["이름", "심볼", "데이터소스"];
      tickerSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row
      tickerSheet.getRange(1, 1, 1, headers.length).setBackground("#f3f3f3");
      tickerSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      
      // Add example rows
      const examples = [
        ["코스피", "KOSPI", "naver"],
        ["코스닥", "KOSDAQ", "naver"],
        ["삼성전자", "005930", "naver"],
        ["S&P 500", "^GSPC", "yahoo"]
      ];
      
      tickerSheet.getRange(2, 1, examples.length, examples[0].length).setValues(examples);
      
      // Freeze header row
      tickerSheet.setFrozenRows(1);
      
      // Adjust column widths
      tickerSheet.setColumnWidth(1, 150);
      tickerSheet.setColumnWidth(2, 100);
      tickerSheet.setColumnWidth(3, 120);
      
      Logger.log(`'${CONFIG.SHEETS.TICKERS}' 시트를 생성하고 샘플 데이터를 추가했습니다.`);
    }
    
    return tickerSheet;
  } catch (error) {
    Logger.log(`티커 시트를 생성하는 중 오류가 발생했습니다: ${error.message}`);
    throw error;
  }
}

/**
 * Get ticker data from the ticker sheet
 * @return {Array} Array of ticker objects
 */
function getTickerData() {
  try {
    const tickerSheet = getOrCreateTickerSheet();
    
    // Make sure we have data on the sheet
    const data = tickerSheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("티커 시트에 데이터가 없습니다.");
      return [];
    }
    
    // Skip header row (first row)
    const tickers = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Make sure the row has all required data
      if (row[0] && row[1] && row[2]) {
        tickers.push({
          name: row[0].trim(),
          ticker: row[1].trim(),
          source: row[2].trim().toLowerCase()  // normalize source names
        });
      }
    }
    
    Logger.log(`${tickers.length}개의 티커를 로드했습니다.`);
    return tickers;
  } catch (error) {
    Logger.log(`티커 데이터를 로드하는 중 오류가 발생했습니다: ${error.message}`);
    throw error;
  }
}

/**
 * Check if a date is a Korean holiday
 * @param {Date} date - The date to check
 * @return {boolean} True if the date is a Korean holiday
 */
function isKoreanHoliday(date) {
  // Format date to YYYY-MM-DD for simple comparison
  const formattedDate = Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd");
  
  // List of Korean holidays for the year (would need to be updated yearly)
  // This includes fixed and lunar holidays for 2023-2025
  const koreanHolidays = [
    // 2023 Korean Holidays
    "2023-01-01", "2023-01-21", "2023-01-22", "2023-01-23", "2023-01-24", // New Year & Seollal
    "2023-03-01", // Independence Movement Day
    "2023-05-05", // Children's Day
    "2023-05-27", // Buddha's Birthday
    "2023-06-06", // Memorial Day
    "2023-08-15", // Liberation Day
    "2023-09-28", "2023-09-29", "2023-09-30", // Chuseok
    "2023-10-03", // National Foundation Day
    "2023-10-09", // Hangul Day
    "2023-12-25", // Christmas
    
    // 2024 Korean Holidays
    "2024-01-01", // New Year's Day
    "2024-02-09", "2024-02-10", "2024-02-11", "2024-02-12", // Seollal
    "2024-03-01", // Independence Movement Day
    "2024-04-10", // Election Day
    "2024-05-05", // Children's Day
    "2024-05-15", // Buddha's Birthday
    "2024-06-06", // Memorial Day
    "2024-08-15", // Liberation Day
    "2024-09-16", "2024-09-17", "2024-09-18", // Chuseok
    "2024-10-03", // National Foundation Day
    "2024-10-09", // Hangul Day
    "2024-12-25", // Christmas
    
    // 2025 Korean Holidays (Preliminary dates)
    "2025-01-01", // New Year's Day
    "2025-01-28", "2025-01-29", "2025-01-30", "2025-01-31", // Seollal
    "2025-03-01", // Independence Movement Day
    "2025-05-05", // Children's Day
    "2025-05-05", // Buddha's Birthday
    "2025-06-06", // Memorial Day
    "2025-08-15", // Liberation Day
    "2025-10-03", // National Foundation Day
    "2025-10-05", "2025-10-06", "2025-10-07", // Chuseok
    "2025-10-09", // Hangul Day
    "2025-12-25"  // Christmas
  ];
  
  return koreanHolidays.includes(formattedDate);
}

/**
 * Get appropriate market memo based on date and time
 * @param {Date} date - The reference date
 * @return {string} Market status memo
 */
function getMarketMemo(date) {
  // Create a copy of the date to use Seoul time
  const seoulDate = new Date(date);
  const seoulOffset = 9; // UTC+9
  seoulDate.setHours(seoulDate.getHours() + seoulOffset - new Date().getTimezoneOffset()/60);
  
  const dayOfWeek = seoulDate.getDay();
  const hours = seoulDate.getHours();
  
  // Check if it's a weekend
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    return "주말 - 한국 시장 휴장. 개장 시간: 월요일 오전 9시";
  }
  
  // Check if it's a Korean holiday
  if (isKoreanHoliday(date)) {
    return "공휴일 - 한국 시장 휴장";
  }
  
  // Check if before market hours (market opens at 9:00 AM KST)
  if (hours < 9) {
    return "장 전 - 한국 시장은 오전 9시에 개장합니다";
  }
  
  // Check if after market hours (market closes at 3:30 PM KST)
  if (hours >= 15 || (hours === 15 && seoulDate.getMinutes() >= 30)) {
    return "장 마감 - 한국 시장은 오전 9시부터 오후 3시 30분까지 개장합니다";
  }
  
  // During market hours
  return "장 중 - 한국 시장 개장 중 (오전 9시 - 오후 3시 30분)";
}

/**
 * Process all tickers and update the dashboard
 * @param {Array} tickers - Array of ticker objects
 * @param {Sheet} dashboardSheet - The dashboard sheet
 * @param {Date} refDate - Reference date
 */
function processTickers(tickers, dashboardSheet, refDate) {
  const dateCalculator = new DateCalculator(refDate);
  const totalTickers = tickers.length;
  
  // Set row height for the header row to make it stand out more
  dashboardSheet.setRowHeight(2, 28);
  
  // Set a consistent row height for all data rows
  for (let i = 3; i < 3 + totalTickers; i++) {
    dashboardSheet.setRowHeight(i, 25);
  }
  
  // Reset progress tracking
  PROGRESS_DATA = {
    percent: 0,
    message: "시작 중...",
    type: "progress",
    tickers: []
  };
  
  // Prepare data array for later formatting
  const data = [];
  
  // First, populate the sheet with placeholders for immediate visual feedback
  const placeholders = tickers.map(ticker => [
    ticker.name,  // Column A - 지수 (Index name)
    "로딩 중...", // Column B - 주간 수익률 (Weekly return)
    "로딩 중...", // Column C - 월간 수익률 (Monthly return)
    "로딩 중...", // Column D - YTD 수익률 (YTD return)
    "로딩 중..."  // Column E - 고점 대비 수익률 (Return vs High)
  ]);
  
  // Render the initial placeholders
  if (placeholders.length > 0) {
    dashboardSheet.getRange(3, 1, placeholders.length, 5).setValues(placeholders);
  }
  
  // Process each ticker individually
  for (let i = 0; i < tickers.length; i++) {
    const ticker = tickers[i];
    try {
      // Update progress
      const progressPercent = Math.round((i / totalTickers) * 100);
      updateProgress(progressPercent, `처리 중 (${i+1}/${totalTickers}): ${ticker.name}`);
      updateLoadingIndicator(`처리 중 (${i+1}/${totalTickers}): ${ticker.name}`, dashboardSheet);
      
      Logger.log(`처리 중: ${ticker.name} (${ticker.ticker}) - 소스: ${ticker.source}`);
      
      // Get prices
      const prices = getPrices(ticker.ticker, ticker.source, dateCalculator);
      
      // Calculate returns
      const returns = calculateReturns(prices);
      
      // Create row data - ensure first column has the ticker name
      const rowData = [
        ticker.name,       // Column A - 지수 (Index name)
        returns.weekly,    // Column B - 주간 수익률 (Weekly return)
        returns.monthly,   // Column C - 월간 수익률 (Monthly return)
        returns.ytd,       // Column D - YTD 수익률 (YTD return)
        returns.high       // Column E - 고점 대비 수익률 (Return vs High)
      ];
      
      // Update this ticker's row in the sheet immediately
      updateTickerRow(i, rowData, dashboardSheet);
      
      // Add to data array for later reference
      data.push(rowData);
      
      // Add to progress data
      PROGRESS_DATA.tickers.push({
        name: ticker.name,
        symbol: ticker.ticker,
        status: "completed"
      });
      
      Logger.log(`처리 완료: ${ticker.name}: ${JSON.stringify(returns)}`);
    } catch (error) {
      Logger.log(`티커 처리 오류 - ${ticker.name}: ${error.message}`);
      
      // Create error row data
      const errorRowData = [
        ticker.name,           // Column A - 지수 (Index name)
        CONFIG.STATUS.ERROR,   // Column B
        CONFIG.STATUS.ERROR,   // Column C
        CONFIG.STATUS.ERROR,   // Column D
        CONFIG.STATUS.ERROR    // Column E
      ];
      
      // Update the row with error indicators immediately
      updateTickerRow(i, errorRowData, dashboardSheet);
      
      // Add to data array for later reference
      data.push(errorRowData);
      
      // Add to progress data
      PROGRESS_DATA.tickers.push({
        name: ticker.name,
        symbol: ticker.ticker,
        status: "error",
        error: error.message
      });
    }
    
    // Force spreadsheet to update after each ticker
    SpreadsheetApp.flush();
    
    // Small delay to allow UI to update
    Utilities.sleep(50);
  }
  
  // Update progress to 100%
  updateProgress(100, "완료");
  
  // Format and add timestamp with more detail
  const now = new Date();
  const timestamp = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  
  // Calculate reference date string
  const referenceDateStr = Utilities.formatDate(refDate, "Asia/Seoul", "yyyy-MM-dd");
  
  // Get market memo based on reference date
  const marketMemo = getMarketMemo(refDate);
  
  // Create a nicely formatted timestamp row
  const timestampRow = data.length + 4;
  const timestampCell = dashboardSheet.getRange("A" + timestampRow);
  timestampCell.setValue(`마지막 업데이트: ${timestamp} (참조일: ${referenceDateStr})`);
  timestampCell.setFontStyle("italic");
  timestampCell.setFontColor("#666666");
  
  // Add market memo
  const memoRow = timestampRow + 1;
  const memoCell = dashboardSheet.getRange("A" + memoRow);
  memoCell.setValue(`📊 ${marketMemo}`);
  memoCell.setFontStyle("italic");
  memoCell.setFontColor("#1a3370");
  memoCell.setBackground("#f5f9ff");
  
  // Set height of the timestamp and memo rows
  dashboardSheet.setRowHeight(timestampRow, 25);
  dashboardSheet.setRowHeight(memoRow, 25);
  
  // Add separator line above the timestamp
  const separatorRange = dashboardSheet.getRange(timestampRow - 1, 1, 1, 5);
  separatorRange.setBorder(false, false, true, false, false, false, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Apply formatting to the dashboard
 * @param {Sheet} dashboardSheet - The dashboard sheet
 */
function formatDashboard(dashboardSheet) {
  // Set optimal column widths for better readability
  dashboardSheet.setColumnWidth(1, 160); // Column A (지수) - wider for index names
  dashboardSheet.setColumnWidth(2, 100); // Column B (주간 수익률)
  dashboardSheet.setColumnWidth(3, 100); // Column C (월간 수익률)
  dashboardSheet.setColumnWidth(4, 100); // Column D (YTD 수익률)
  dashboardSheet.setColumnWidth(5, 120); // Column E (고점 대비 수익률) - slightly wider
  
  // Style the header row
  const headerRange = dashboardSheet.getRange("A2:E2");
  headerRange.setBackground("#f3f3f3");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  headerRange.setBorder(true, true, true, true, false, false);
  
  // Set alignment for data columns
  const lastRow = Math.max(dashboardSheet.getLastRow(), 3);
  if (lastRow > 2) {
    // Center align return columns
    dashboardSheet.getRange(3, 2, lastRow - 2, 4).setHorizontalAlignment("center");
    
    // Left align ticker names
    dashboardSheet.getRange(3, 1, lastRow - 2, 1).setHorizontalAlignment("left");
    
    // Add zebra striping for better readability
    for (let i = 3; i <= lastRow; i++) {
      if ((i - 3) % 2 === 0) {
        dashboardSheet.getRange(i, 1, 1, 5).setBackground("#f9f9f9");
      }
    }
    
    // Add bottom border to the last data row
    dashboardSheet.getRange(lastRow, 1, 1, 5).setBorder(false, false, true, false, false, false);
  }
  
  // Format timestamp row if it exists
  if (dashboardSheet.getRange("A" + (lastRow + 1)).getValue().includes("마지막 업데이트")) {
    dashboardSheet.getRange("A" + (lastRow + 1)).setFontStyle("italic");
    dashboardSheet.getRange("A" + (lastRow + 1)).setFontColor("#666666");
  }
}

/**
 * Get the last Korean business day
 * @param {Date} [date] - Optional date to calculate from, defaults to today
 * @return {Date} The last business day
 */
function getLastKoreanBusinessDay(date) {
  // If no date is provided, use today
  const inputDate = date || new Date();
  
  // Create a copy of the date to work with
  const adjustedDate = new Date(inputDate);
  
  // Set to Korea timezone (for proper day comparison)
  const koreaTime = new Date(inputDate.toLocaleString("en-US", {timeZone: "Asia/Seoul"}));
  const dayOfWeek = koreaTime.getDay();
  const hours = koreaTime.getHours();
  
  // If it's before market open and a weekday, use previous day
  if (hours < 9 && dayOfWeek >= 1 && dayOfWeek <= 5) {
    adjustedDate.setDate(adjustedDate.getDate() - 1);
  }
  
  // Keep going back until we find a business day
  let attempts = 0; // Safety counter
  while ((adjustedDate.getDay() === 0 || 
          adjustedDate.getDay() === 6 || 
          isKoreanHoliday(adjustedDate)) && attempts < 10) {
    adjustedDate.setDate(adjustedDate.getDate() - 1);
    attempts++;
  }
  
  // Log the adjustment if it's different from the input date
  if (!datesEqual(inputDate, adjustedDate)) {
    const originalDateStr = formatDate(inputDate);
    const adjustedDateStr = formatDate(adjustedDate);
    Logger.log(`참조일 조정: ${originalDateStr} → ${adjustedDateStr} (마지막 영업일 기준)`);
  }
  
  return adjustedDate;
}

/**
 * Compare if two dates are the same (ignoring time)
 * @param {Date} date1 - First date
 * @param {Date} date2 - Second date
 * @return {boolean} True if dates are the same
 */
function datesEqual(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

/**
 * Format a date as YYYY-MM-DD
 * @param {Date} date - The date to format
 * @return {string} Formatted date string
 */
function formatDate(date) {
  return Utilities.formatDate(date, "Asia/Seoul", "yyyy-MM-dd");
}

/**
 * Get the dashboard sheet or create it if it doesn't exist
 * @return {Sheet} The dashboard sheet
 */
function getDashboardSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = CONFIG && CONFIG.SHEETS && CONFIG.SHEETS.DASHBOARD 
      ? CONFIG.SHEETS.DASHBOARD 
      : "dashboard";
    
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      // Create the dashboard sheet if it doesn't exist
      sheet = ss.insertSheet(sheetName);
      Logger.log(`'${sheetName}' 시트를 생성했습니다.`);
      
      // Set up initial formatting for new sheet
      sheet.setColumnWidth(1, 160); // Name column
      sheet.setFrozenRows(2);       // Freeze header rows
    }
    
    return sheet;
  } catch (error) {
    Logger.log(`대시보드 시트를 생성하는 중 오류가 발생했습니다: ${error.message}`);
    throw error;
  }
}

/**
 * Get or create the temporary calculation sheet
 * @return {Sheet} The temporary calculation sheet
 */
function getOrCreateTempSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let tempSheet = ss.getSheetByName(CONFIG.SHEETS.TEMP);
    
    if (!tempSheet) {
      Logger.log("임시 계산 시트를 생성합니다.");
      tempSheet = ss.insertSheet(CONFIG.SHEETS.TEMP);
      
      // Configure the sheet
      tempSheet.hideSheet();
      
      // Add a header to identify the sheet's purpose
      tempSheet.getRange("A1:D1").merge();
      tempSheet.getRange("A1").setValue("임시 계산 영역 - 편집하지 마세요!");
      tempSheet.getRange("A1").setFontWeight("bold");
      tempSheet.getRange("A1").setBackground("#ffeeee");
      
      // Label the sections
      tempSheet.getRange("A2").setValue("Google Price");
      tempSheet.getRange("B2").setValue("Google High");
      tempSheet.getRange("C2").setValue("YTD Date");
      tempSheet.getRange("D2").setValue("General Use");
      
      // Set column widths
      tempSheet.setColumnWidth(1, 150);
      tempSheet.setColumnWidth(2, 150);
      tempSheet.setColumnWidth(3, 150);
      tempSheet.setColumnWidth(4, 200);
    }
    
    return tempSheet;
  } catch (error) {
    Logger.log(`임시 계산 시트 생성 중 오류: ${error.message}`);
    return null;
  }
}

/**
 * Initialize the application
 * Called when the spreadsheet is opened
 */
function initialize() {
  try {
    // Ensure temporary calculation sheet exists
    getOrCreateTempSheet();
    
    // Set up custom menu
    onOpen();
    
    Logger.log("애플리케이션이 초기화되었습니다.");
  } catch (error) {
    Logger.log(`초기화 중 오류 발생: ${error.message}`);
  }
}

/**
 * Get all required prices for a given ticker
 * @param {string} symbol - Ticker symbol
 * @param {string} source - Data source 
 * @param {DateCalculator} dateCalculator - Date calculator instance
 * @return {Object} Object containing all relevant prices
 */
function getPrices(symbol, source, dateCalculator) {
  try {
    Logger.log(`가격 데이터 가져오기 시작 - 심볼: ${symbol}, 소스: ${source}`);
    Logger.log(`참조일: ${dateCalculator.referenceDate}`);
    Logger.log(`주간 기준일: ${dateCalculator.getWeekAgoDate()}`);
    Logger.log(`월간 기준일: ${dateCalculator.getMonthAgoDate()}`);
    Logger.log(`연초 기준일: ${dateCalculator.getYtdDate()}`);
    
    const dataProvider = new DataProviderFactory().getProvider(source);
    Logger.log(`데이터 제공자: ${dataProvider.name || source}`);
    
    // Format the symbol appropriately for the data provider
    const formattedSymbol = DataProviderFactory.formatSymbol(symbol, source);
    
    // Log the symbol conversion for debugging
    if (formattedSymbol !== symbol) {
      Logger.log(`심볼 변환: ${symbol} -> ${formattedSymbol}`);
    }
    
    // Get current price
    Logger.log(`현재가 요청 시작 - 심볼: ${formattedSymbol}, 참조일: ${dateCalculator.referenceDate}`);
    const currentPrice = dataProvider.getPrice(formattedSymbol, dateCalculator.referenceDate);
    Logger.log(`현재가: ${currentPrice}`);
    
    // Get week ago price
    Logger.log(`주간 가격 요청 시작 - 심볼: ${formattedSymbol}, 참조일: ${dateCalculator.getWeekAgoDate()}`);
    const weekAgoPrice = dataProvider.getPrice(formattedSymbol, dateCalculator.getWeekAgoDate());
    Logger.log(`주간 가격: ${weekAgoPrice}`);
    
    // Get month ago price
    Logger.log(`월간 가격 요청 시작 - 심볼: ${formattedSymbol}, 참조일: ${dateCalculator.getMonthAgoDate()}`);
    const monthAgoPrice = dataProvider.getPrice(formattedSymbol, dateCalculator.getMonthAgoDate());
    Logger.log(`월간 가격: ${monthAgoPrice}`);
    
    // Get YTD price
    Logger.log(`연초대비 가격 요청 시작 - 심볼: ${formattedSymbol}, 참조일: ${dateCalculator.getYtdDate()}`);
    const ytdPrice = dataProvider.getPrice(formattedSymbol, dateCalculator.getYtdDate());
    Logger.log(`연초 가격: ${ytdPrice}`);
    
    // Get 52-week high price
    Logger.log(`52주 고점 요청 시작 - 심볼: ${formattedSymbol}`);
    const highPrice = dataProvider.getHighPrice(formattedSymbol);
    Logger.log(`52주 고점: ${highPrice}`);
    
    const prices = {
      current: currentPrice,
      weekAgo: weekAgoPrice,
      monthAgo: monthAgoPrice,
      ytd: ytdPrice,
      high: highPrice
    };
    
    Logger.log(`가격 데이터 수집 완료 - 심볼: ${symbol}`);
    Logger.log(JSON.stringify(prices));
    
    return prices;
  } catch (error) {
    Logger.log(`가격 데이터 가져오기 오류 - 심볼: ${symbol}, 오류: ${error.message}, 스택: ${error.stack}`);
    throw error;
  }
}