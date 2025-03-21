/**
 * Performance Dashboard Ticker - Diagnostics Service Module
 * 
 * This module provides diagnostic tools for troubleshooting issues with the dashboard.
 */

/**
 * Run diagnostics for a specific symbol
 * @param {string} symbol - Symbol to diagnose
 * @param {string} source - Data source for the symbol
 * @return {Object} Diagnostic results
 */
function runSymbolDiagnostics(symbol, source) {
  try {
    Logger.log(`심볼 진단 시작: ${symbol} (${source})`);
    
    // Create temp sheet for diagnostics if it doesn't exist
    let tempSheet = SS.getSheetByName(CONFIG.SHEETS.TEMP);
    if (!tempSheet) {
      tempSheet = SS.insertSheet(CONFIG.SHEETS.TEMP);
    }
    
    // Clear the temp sheet
    tempSheet.clear();
    
    // Set headers
    tempSheet.getRange(1, 1, 1, 4).setValues([["진단 항목", "결과", "기대값", "비고"]]);
    tempSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    tempSheet.setFrozenRows(1);
    
    // Initialize row counter for results
    let rowCounter = 2;
    
    // Initialize diagnostic results
    const diagnosticResults = {
      symbol: symbol,
      source: source,
      tests: [],
      success: true,
      timestamp: new Date().toISOString()
    };
    
    // Test symbol format
    try {
      const dataProviderFactory = new DataProviderFactory(new MarketTimeManager());
      const provider = dataProviderFactory.getProvider(source);
      const formattedSymbol = provider.formatSymbol(symbol);
      
      tempSheet.getRange(rowCounter, 1, 1, 4).setValues([
        ["심볼 포맷", formattedSymbol, symbol, "소스별 포맷팅 적용"]
      ]);
      
      diagnosticResults.tests.push({
        name: "Symbol Format",
        result: "PASS",
        value: formattedSymbol
      });
      
      rowCounter++;
    } catch (error) {
      tempSheet.getRange(rowCounter, 1, 1, 4).setValues([
        ["심볼 포맷", "오류", symbol, error.message]
      ]);
      
      diagnosticResults.tests.push({
        name: "Symbol Format",
        result: "FAIL",
        error: error.message
      });
      
      diagnosticResults.success = false;
      rowCounter++;
    }
    
    // Test current price retrieval
    try {
      const dataProviderFactory = new DataProviderFactory(new MarketTimeManager());
      const provider = dataProviderFactory.getProvider(source);
      const currentPrice = provider.getPrice(symbol, new Date());
      
      tempSheet.getRange(rowCounter, 1, 1, 4).setValues([
        ["현재 가격", currentPrice, "숫자값", "현재 가격 조회"]
      ]);
      
      diagnosticResults.tests.push({
        name: "Current Price",
        result: "PASS",
        value: currentPrice
      });
      
      rowCounter++;
    } catch (error) {
      tempSheet.getRange(rowCounter, 1, 1, 4).setValues([
        ["현재 가격", "오류", "숫자값", error.message]
      ]);
      
      diagnosticResults.tests.push({
        name: "Current Price",
        result: "FAIL",
        error: error.message
      });
      
      diagnosticResults.success = false;
      rowCounter++;
    }
    
    // Test historical price retrieval
    try {
      const dataProviderFactory = new DataProviderFactory(new MarketTimeManager());
      const provider = dataProviderFactory.getProvider(source);
      const dateCalculator = new DateCalculator(new Date());
      const weeklyDate = dateCalculator.getDateWeeksAgo(1);
      const historicalPrice = provider.getHistoricalPrice(symbol, weeklyDate);
      
      tempSheet.getRange(rowCounter, 1, 1, 4).setValues([
        ["1주일 전 가격", historicalPrice, "숫자값", `${weeklyDate.toISOString()} 기준 가격`]
      ]);
      
      diagnosticResults.tests.push({
        name: "Historical Price",
        result: "PASS",
        value: historicalPrice,
        date: weeklyDate.toISOString()
      });
      
      rowCounter++;
    } catch (error) {
      tempSheet.getRange(rowCounter, 1, 1, 4).setValues([
        ["1주일 전 가격", "오류", "숫자값", error.message]
      ]);
      
      diagnosticResults.tests.push({
        name: "Historical Price",
        result: "FAIL",
        error: error.message
      });
      
      diagnosticResults.success = false;
      rowCounter++;
    }
    
    // Test 52-week high price retrieval
    try {
      const dataProviderFactory = new DataProviderFactory(new MarketTimeManager());
      const provider = dataProviderFactory.getProvider(source);
      const highPrice = provider.getHighPrice(symbol);
      
      tempSheet.getRange(rowCounter, 1, 1, 4).setValues([
        ["52주 최고가", highPrice, "숫자값", "52주 최고가 조회"]
      ]);
      
      diagnosticResults.tests.push({
        name: "52-Week High",
        result: "PASS",
        value: highPrice
      });
      
      rowCounter++;
    } catch (error) {
      tempSheet.getRange(rowCounter, 1, 1, 4).setValues([
        ["52주 최고가", "오류", "숫자값", error.message]
      ]);
      
      diagnosticResults.tests.push({
        name: "52-Week High",
        result: "FAIL",
        error: error.message
      });
      
      diagnosticResults.success = false;
      rowCounter++;
    }
    
    // Auto-resize columns
    tempSheet.autoResizeColumns(1, 4);
    
    // Make the sheet visible
    tempSheet.showSheet();
    
    // Set active sheet to the temp sheet
    SS.setActiveSheet(tempSheet);
    
    Logger.log(`심볼 진단 완료: ${symbol} (${source})`);
    return diagnosticResults;
  } catch (error) {
    Logger.log(`심볼 진단 오류: ${error.message}`);
    return {
      symbol: symbol,
      source: source,
      tests: [],
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Show diagnostic sheet
 */
function showTempSheet() {
  try {
    let tempSheet = SS.getSheetByName(CONFIG.SHEETS.TEMP);
    
    if (!tempSheet) {
      tempSheet = SS.insertSheet(CONFIG.SHEETS.TEMP);
      tempSheet.getRange(1, 1).setValue("진단 도구를 사용하여 데이터를 생성하세요.");
    }
    
    // Make the sheet visible and active
    tempSheet.showSheet();
    SS.setActiveSheet(tempSheet);
    
    Logger.log('진단 시트가 표시되었습니다.');
  } catch (error) {
    Logger.log(`진단 시트 표시 오류: ${error.message}`);
    showErrorAlert('진단 시트 표시 오류', error.message);
  }
}

/**
 * Show symbol diagnostics dialog
 */
function showSymbolDiagnosticsDialog() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      '🔍 심볼 진단',
      '진단할 심볼과 데이터 소스를 입력하세요. (형식: SYMBOL,SOURCE)\n' +
      '예: 005930,naver 또는 AAPL,google',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    const inputText = response.getResponseText().trim();
    const parts = inputText.split(',');
    
    if (parts.length !== 2) {
      showAlert('⚠️ 잘못된 입력', '올바른 형식으로 입력하세요. (예: AAPL,google)');
      return;
    }
    
    const symbol = parts[0].trim();
    const source = parts[1].trim().toLowerCase();
    
    if (!symbol) {
      showAlert('⚠️ 잘못된 입력', '심볼을 입력하세요.');
      return;
    }
    
    if (!['google', 'yahoo', 'naver'].includes(source)) {
      showAlert('⚠️ 잘못된 입력', '유효한 데이터 소스를 입력하세요. (google, yahoo, naver)');
      return;
    }
    
    // Run diagnostics
    showAlert('🔍 진단 시작', `${symbol} (${source}) 진단을 시작합니다. 완료되면 진단 시트가 표시됩니다.`);
    const results = runSymbolDiagnostics(symbol, source);
    
    if (results.success) {
      showAlert('✅ 진단 완료', `${symbol} (${source}) 진단이 완료되었습니다. 진단 시트를 확인하세요.`);
    } else {
      showAlert('⚠️ 진단 실패', `${symbol} (${source}) 진단 중 오류가 발생했습니다. 진단 시트를 확인하세요.`);
    }
  } catch (error) {
    Logger.log(`심볼 진단 대화 상자 오류: ${error.message}`);
    showErrorAlert('심볼 진단 오류', error.message);
  }
}

/**
 * Force reset all system settings
 * This is an emergency function for when the system is stuck
 */
function forceResetSystem() {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '⚠️ 시스템 강제 초기화',
      '정말로 모든 시스템 설정을 초기화하시겠습니까?\n\n' +
      '이 작업은 모든 잠금 및 진행 상태를 초기화합니다.',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // Reset all script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteAllProperties();
    
    // Reset progress tracking
    PROGRESS_DATA = {
      percent: 0,
      message: "",
      type: "ready",
      tickers: []
    };
    
    // Force unlock
    const lockStatus = forceResetLock();
    
    // Show confirmation
    showAlert(
      '✅ 시스템 초기화 완료',
      '모든 시스템 설정이 초기화되었습니다.\n\n' +
      '이제 대시보드 업데이트를 다시 시도할 수 있습니다.'
    );
    
    Logger.log('시스템 강제 초기화 완료');
  } catch (error) {
    Logger.log(`시스템 초기화 오류: ${error.message}`);
    showErrorAlert('시스템 초기화 오류', error.message);
  }
}

/**
 * Generate a debug report for the user to share
 * This function is called from the UI
 */
function generateDebugReport() {
  return createDebugReport();
} 