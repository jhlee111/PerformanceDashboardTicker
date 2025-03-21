/**
 * Performance Dashboard Ticker - Dashboard Service Module
 * 
 * This module provides functionality for managing the dashboard.
 */

/**
 * Render data to the dashboard
 * @param {Sheet} sheet - The dashboard sheet
 * @param {number} row - The row to update
 * @param {Object} tickerData - The ticker data to render
 * @param {Object} columnMap - Mapping of data to column indices
 */
function renderToDashboard(sheet, row, tickerData, columnMap) {
  try {
    const { name, ticker, source, prices, returns, dateAnnotation } = tickerData;
    
    // Set ticker information
    sheet.getRange(row, columnMap.NAME).setValue(name);
    sheet.getRange(row, columnMap.TICKER).setValue(ticker);
    sheet.getRange(row, columnMap.SOURCE).setValue(source.toUpperCase());
    
    // Set price values
    if (prices) {
      sheet.getRange(row, columnMap.CURRENT).setValue(prices.current);
      sheet.getRange(row, columnMap.HIGH).setValue(prices.high);
    }
    
    // Set return values and apply formatting
    if (returns) {
      if (returns.weekly !== null) {
        sheet.getRange(row, columnMap.WEEKLY).setValue(returns.weekly / 100); // Convert to decimal for percentage formatting
        formatReturnCell(sheet, row, columnMap.WEEKLY, returns.weekly / 100);
      } else {
        sheet.getRange(row, columnMap.WEEKLY).setValue('N/A');
      }
      
      if (returns.monthly !== null) {
        sheet.getRange(row, columnMap.MONTHLY).setValue(returns.monthly / 100);
        formatReturnCell(sheet, row, columnMap.MONTHLY, returns.monthly / 100);
      } else {
        sheet.getRange(row, columnMap.MONTHLY).setValue('N/A');
      }
      
      if (returns.ytd !== null) {
        sheet.getRange(row, columnMap.YTD).setValue(returns.ytd / 100);
        formatReturnCell(sheet, row, columnMap.YTD, returns.ytd / 100);
      } else {
        sheet.getRange(row, columnMap.YTD).setValue('N/A');
      }
      
      if (returns.high !== null) {
        sheet.getRange(row, columnMap.HIGH_RETURN).setValue(returns.high / 100);
        formatReturnCell(sheet, row, columnMap.HIGH_RETURN, returns.high / 100);
      } else {
        sheet.getRange(row, columnMap.HIGH_RETURN).setValue('N/A');
      }
    }
    
    // Set date annotation if available
    if (dateAnnotation) {
      sheet.getRange(row, columnMap.INFO).setValue(dateAnnotation);
      sheet.getRange(row, columnMap.INFO).setFontColor("#1565C0"); // Blue for information
    }
    
    // Set last updated timestamp
    const now = new Date();
    const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(row, columnMap.LAST_UPDATED).setValue(formattedDate);
    
    Logger.log(`Rendered data for ${name} (${ticker}) to row ${row}`);
  } catch (error) {
    Logger.log(`대시보드 렌더링 오류 (${tickerData.name}): ${error.message}`);
    showTickerError(sheet, row, columnMap.INFO, error.message);
  }
}

/**
 * Initialize the dashboard headers and formatting
 * @param {Sheet} sheet - The dashboard sheet
 * @return {Object} Column mapping for the dashboard
 */
function initializeDashboard(sheet) {
  try {
    Logger.log('대시보드 초기화 중...');
    
    // Define headers
    const headers = [
      '이름', '티커', '소스', '현재가', '52주 최고가', '주간 변화', '월간 변화', 
      'YTD 변화', '최고가 대비', '정보', '마지막 업데이트'
    ];
    
    // Get reference date for display
    const referenceDate = getReferenceDate();
    const formattedReferenceDate = Utilities.formatDate(referenceDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    // Set headers in row 1
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    
    // Format headers
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#E3F2FD'); // Light blue background
    
    // Add reference date to "정보" column header
    const infoHeaderCell = sheet.getRange(1, 10); // "정보" column
    infoHeaderCell.setValue('정보\n(기준일: ' + formattedReferenceDate + ')');
    infoHeaderCell.setFontWeight('bold');
    infoHeaderCell.setFontColor('#1a73e8');
    infoHeaderCell.setWrap(true);
    infoHeaderCell.setVerticalAlignment('middle');
    
    // Freeze the header row
    sheet.setFrozenRows(1);
    
    // Create column mapping for easier reference
    const columnMap = {
      NAME: 1,
      TICKER: 2,
      SOURCE: 3,
      CURRENT: 4,
      HIGH: 5,
      WEEKLY: 6,
      MONTHLY: 7,
      YTD: 8,
      HIGH_RETURN: 9,
      INFO: 10,
      LAST_UPDATED: 11
    };
    
    // Format price columns (start from row 2 now)
    sheet.getRange(2, columnMap.CURRENT, sheet.getMaxRows() - 1, 1).setNumberFormat('#,##0.00');
    sheet.getRange(2, columnMap.HIGH, sheet.getMaxRows() - 1, 1).setNumberFormat('#,##0.00');
    
    // Format return columns
    sheet.getRange(2, columnMap.WEEKLY, sheet.getMaxRows() - 1, 1).setNumberFormat('0.00%');
    sheet.getRange(2, columnMap.MONTHLY, sheet.getMaxRows() - 1, 1).setNumberFormat('0.00%');
    sheet.getRange(2, columnMap.YTD, sheet.getMaxRows() - 1, 1).setNumberFormat('0.00%');
    sheet.getRange(2, columnMap.HIGH_RETURN, sheet.getMaxRows() - 1, 1).setNumberFormat('0.00%');
    
    // Set appropriate column widths based on header content
    const headerWidths = {
      [columnMap.NAME]: 120,       // 이름
      [columnMap.TICKER]: 180,     // 티커
      [columnMap.SOURCE]: 80,      // 소스
      [columnMap.CURRENT]: 90,     // 현재가
      [columnMap.HIGH]: 100,       // 52주 최고가
      [columnMap.WEEKLY]: 80,      // 주간 변화
      [columnMap.MONTHLY]: 80,     // 월간 변화
      [columnMap.YTD]: 80,         // YTD 변화
      [columnMap.HIGH_RETURN]: 90, // 최고가 대비
      [columnMap.INFO]: 150,       // 정보 (including reference date)
      [columnMap.LAST_UPDATED]: 140 // 마지막 업데이트
    };
    
    // Apply column widths from the predefined widths
    for (const [column, width] of Object.entries(headerWidths)) {
      sheet.setColumnWidth(parseInt(column), width);
    }
    
    // Increase the height of the header row to accommodate the reference date info
    sheet.setRowHeight(1, 50);
    
    Logger.log('대시보드 초기화 완료');
    
    return columnMap;
  } catch (error) {
    Logger.log(`대시보드 초기화 오류: ${error.message}`);
    throw new Error(`대시보드를 초기화할 수 없습니다: ${error.message}`);
  }
}

/**
 * Clear existing data from the dashboard
 * @param {Sheet} sheet - The dashboard sheet
 */
function clearDashboardData(sheet) {
  try {
    Logger.log('기존 대시보드 데이터 지우는 중...');
    
    // Get the number of rows with data
    const lastRow = sheet.getLastRow();
    
    // If there's data beyond the header row, clear it
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
      dataRange.clearContent();
      dataRange.clearFormat();
      
      Logger.log(`${lastRow - 1}개 행의 데이터를 지웠습니다.`);
    } else {
      Logger.log('지울 데이터가 없습니다.');
    }
    
    // Don't need to preserve reference date display row anymore since it's part of the header
  } catch (error) {
    Logger.log(`대시보드 데이터 지우기 오류: ${error.message}`);
  }
}

/**
 * Update performance dashboard with ticker data
 */
function updatePerformanceDashboard() {
  try {
    // Check if update is already in progress
    if (isUpdateInProgress()) {
      const lockStatus = getLockStatus();
      
      if (lockStatus.locked) {
        const formattedTime = Utilities.formatDate(
          lockStatus.time, 
          Session.getScriptTimeZone(), 
          'yyyy-MM-dd HH:mm:ss'
        );
        
        showAlert(
          '⚠️ 업데이트 이미 진행 중',
          `다른 사용자(${lockStatus.user})가 이미 ${formattedTime}에 업데이트를 시작했습니다.\n\n` +
          '업데이트가 완료될 때까지 기다려 주세요. 5분 이상 경과된 경우 "잠금 강제 해제" 메뉴를 사용할 수 있습니다.'
        );
      }
      
      return;
    }
    
    // Acquire lock
    if (!acquireUpdateLock()) {
      return;
    }
    
    Logger.log('성능 대시보드 업데이트 시작...');
    
    // Clear audit data - always enable audit for each update
    clearAuditData();
    
    // Get or create dashboard sheet
    const sheet = getDashboardSheet();
    
    // Initialize dashboard and get column mapping
    clearDashboardData(sheet);
    const columnMap = initializeDashboard(sheet);
    
    // Get reference date
    const referenceDate = getReferenceDate();
    Logger.log(`기준일: ${referenceDate.toISOString()}`);
    
    // Create date calculator
    const dateCalculator = new DateCalculator(referenceDate);
    
    // Load ticker data
    const tickers = loadTickerData();
    
    if (tickers.length === 0) {
      showAlert('⚠️ 티커 없음', '처리할 티커가 없습니다. 티커 시트에 데이터를 추가하세요.');
      clearUpdateLock();
      
      // Make sure to set the dashboard as the active sheet
      SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
      
      return;
    }
    
    // Process each ticker and update dashboard - start at row 2 now (after header)
    for (let i = 0; i < tickers.length; i++) {
      try {
        const ticker = tickers[i];
        const row = i + 2; // Start at row 2 (after header)
        
        // Show loading indicator
        showLoadingIndicator(sheet, row, columnMap.INFO, `${ticker.name} 로딩 중...`);
        
        // Process ticker
        Logger.log(`티커 처리 중 (${i + 1}/${tickers.length}): ${ticker.name} (${ticker.ticker})`);
        const processedData = processTicker(ticker, dateCalculator);
        
        // Render to dashboard
        renderToDashboard(sheet, row, processedData, columnMap);
        
        // Hide loading indicator
        hideLoadingIndicator(sheet, row, columnMap.INFO);
      } catch (error) {
        Logger.log(`티커 처리 오류 (${tickers[i].name}): ${error.message}`);
        showTickerError(sheet, i + 2, columnMap.INFO, error.message); // Error at row 2+
      }
    }
    
    // Update complete, release lock
    clearUpdateLock();
    
    // Make sure the dashboard is the active sheet at the end of the update
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
    
    Logger.log('성능 대시보드 업데이트 완료');
    showAlert('✅ 업데이트 완료', `${tickers.length}개의 티커를 성공적으로 업데이트했습니다.\n\n감사 데이터도 업데이트되었습니다.`);
    
    // Return true to indicate successful update
    return true;
  } catch (error) {
    Logger.log(`대시보드 업데이트 오류: ${error.message}`);
    showErrorAlert('대시보드 업데이트 실패', error.message);
    
    // Make sure to release lock even if error occurs
    clearUpdateLock();
    
    // Make sure the dashboard is the active sheet even if there's an error
    try {
      const sheet = getDashboardSheet();
      SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
    } catch (e) {
      // Ignore errors when trying to set active sheet during error handling
      Logger.log(`대시보드 활성화 실패: ${e.message}`);
    }
    
    // Return false to indicate failed update
    return false;
  }
}

/**
 * Toggle diagnostic mode
 */
function toggleDiagnosticMode() {
  try {
    const currentStatus = isDiagnosticModeEnabled();
    enableDiagnosticMode(!currentStatus);
    
    showAlert(
      '🔬 진단 모드',
      `진단 모드가 ${!currentStatus ? '활성화' : '비활성화'}되었습니다.`
    );
  } catch (error) {
    Logger.log(`진단 모드 전환 오류: ${error.message}`);
    showErrorAlert('진단 모드 전환 실패', error.message);
  }
} 