/**
 * Performance Dashboard Ticker - Spreadsheet Utilities Module
 */

/**
 * Get the dashboard sheet
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