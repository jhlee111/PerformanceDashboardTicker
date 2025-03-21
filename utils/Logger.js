/**
 * Performance Dashboard Ticker - Logger Module
 * 
 * Enhanced logging utilities for the application.
 */

/**
 * Key for diagnostic mode in script properties
 */
const DIAGNOSTIC_MODE_KEY = 'diagnosticMode';

/**
 * Log debugging information (only in diagnostic mode)
 * @param {string} message - The message to log
 * @param {*} data - Optional data to log
 */
function debug(message, data) {
  if (!isDiagnosticModeEnabled()) {
    return;
  }
  
  if (data) {
    Logger.log(`[DEBUG] ${message}`);
    Logger.log(JSON.stringify(data, null, 2));
  } else {
    Logger.log(`[DEBUG] ${message}`);
  }
}

/**
 * Log information
 * @param {string} message - The message to log
 */
function info(message) {
  Logger.log(`[INFO] ${message}`);
}

/**
 * Log an error
 * @param {string} message - The error message
 * @param {Error} error - The error object
 */
function error(message, error) {
  if (error) {
    Logger.log(`[ERROR] ${message}: ${error.message}`);
    if (error.stack) {
      Logger.log(`[STACK] ${error.stack}`);
    }
  } else {
    Logger.log(`[ERROR] ${message}`);
  }
}

/**
 * Enable or disable diagnostic mode
 * @param {boolean} enable - True to enable, false to disable
 * @param {boolean} showAlert - Whether to show an alert (default: true)
 * @return {boolean} True if successful
 */
function enableDiagnosticMode(enable = true, showAlert = true) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(DIAGNOSTIC_MODE_KEY, enable ? 'true' : 'false');
  
  if (showAlert) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      enable ? '🔬 진단 모드 활성화' : '🔬 진단 모드 비활성화',
      enable ? 
        '진단 모드가 활성화되었습니다. 로그에 추가 디버깅 정보가 포함됩니다.' : 
        '진단 모드가 비활성화되었습니다.',
      ui.ButtonSet.OK
    );
  }
  
  const statusMessage = enable ? '진단 모드가 활성화되었습니다.' : '진단 모드가 비활성화되었습니다.';
  
  Logger.log(`[SYSTEM] ${statusMessage}`);
  return true;
}

/**
 * Check if diagnostic mode is enabled
 * @return {boolean} True if diagnostic mode is enabled
 */
function isDiagnosticModeEnabled() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty(DIAGNOSTIC_MODE_KEY) === 'true';
}

/**
 * Log price data in a formatted way
 * @param {string} description - Description of the price
 * @param {number} price - The price value
 * @param {string} symbol - The ticker symbol
 */
function logPrice(description, price, symbol) {
  if (!isDiagnosticModeEnabled()) return;
  
  try {
    const formattedPrice = typeof price === 'number' 
      ? price.toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2})
      : price;
    
    Logger.log(`[PRICE] ${symbol} ${description}: ${formattedPrice}`);
  } catch (error) {
    // Silently fail for price logs
  }
}

/**
 * Create a debug report in the temp sheet
 * This will create a formatted report of logs and data that can be copied and shared
 */
function createDebugReport() {
  try {
    // Get or create the temp sheet
    let tempSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.TEMP);
    if (!tempSheet) {
      tempSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(CONFIG.SHEETS.TEMP);
    }
    
    // Clear the sheet
    tempSheet.clear();
    
    // Add title and timestamp
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    tempSheet.getRange(1, 1).setValue('성능 대시보드 디버그 보고서');
    tempSheet.getRange(2, 1).setValue(`생성 시간: ${timestamp}`);
    
    // Format the header
    tempSheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);
    tempSheet.getRange(2, 1).setFontStyle('italic');
    
    // Add system information
    tempSheet.getRange(4, 1).setValue('시스템 정보:');
    tempSheet.getRange(4, 1).setFontWeight('bold');
    
    const systemInfo = [
      ['스크립트 ID', ScriptApp.getScriptId()],
      ['타임존', Session.getScriptTimeZone()],
      ['사용자', Session.getActiveUser().getEmail()],
      ['진단 모드', isDiagnosticModeEnabled() ? '활성화' : '비활성화'],
      ['현재 날짜', Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd')],
      ['현재 시간', Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss')]
    ];
    
    tempSheet.getRange(5, 1, systemInfo.length, 2).setValues(systemInfo);
    
    // Get recent logs
    const logs = Logger.getLog().split('\n');
    
    // Add logs section
    tempSheet.getRange(5 + systemInfo.length + 2, 1).setValue('최근 로그:');
    tempSheet.getRange(5 + systemInfo.length + 2, 1).setFontWeight('bold');
    
    // Filter out empty log entries
    const filteredLogs = logs.filter(log => log.trim().length > 0);
    
    // Add each log line
    for (let i = 0; i < filteredLogs.length; i++) {
      tempSheet.getRange(5 + systemInfo.length + 3 + i, 1).setValue(filteredLogs[i]);
      
      // Color-code by log type
      if (filteredLogs[i].includes('[ERROR]')) {
        tempSheet.getRange(5 + systemInfo.length + 3 + i, 1).setFontColor('#D32F2F');
      } else if (filteredLogs[i].includes('[DEBUG]')) {
        tempSheet.getRange(5 + systemInfo.length + 3 + i, 1).setFontColor('#0D47A1');
      } else if (filteredLogs[i].includes('[PRICE]')) {
        tempSheet.getRange(5 + systemInfo.length + 3 + i, 1).setFontColor('#2E7D32');
      }
    }
    
    // Add a note about how to share
    tempSheet.getRange(5 + systemInfo.length + 3 + filteredLogs.length + 2, 1).setValue(
      '이 보고서를 공유하려면, 시트를 선택하고 Ctrl+A (또는 Cmd+A), Ctrl+C (또는 Cmd+C)로 복사한 후 메시지에 붙여넣기 하세요.'
    );
    
    // Format the sheet
    tempSheet.autoResizeColumns(1, 2);
    
    // Activate the sheet
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(tempSheet);
    
    // Return success message
    return {
      success: true,
      message: '디버그 보고서가 생성되었습니다. "TempCalc" 시트에서 확인하세요.'
    };
  } catch (error) {
    Logger.log(`디버그 보고서 생성 오류: ${error.message}`);
    return {
      success: false,
      message: `디버그 보고서 생성 중 오류가 발생했습니다: ${error.message}`
    };
  }
} 