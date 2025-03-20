/**
 * Provides functionality for retrieving and parsing ticker data from the spreadsheet
 */

/**
 * Get the list of tickers from the tickers sheet
 * @return {Array} Array of ticker objects
 */
function getTickerList() {
  try {
    // Get ticker sheet
    const tickerSheet = SS.getSheetByName(CONFIG.SHEETS.TICKERS);
    if (!tickerSheet) {
      Logger.log("티커 시트를 찾을 수 없습니다.");
      return [];
    }
    
    // Instead of getDataRange(), which includes empty rows,
    // get only the rows that actually contain data
    const lastRow = tickerSheet.getLastRow();
    const lastCol = tickerSheet.getLastColumn();
    
    // If sheet is empty or only has a header row, return empty array
    if (lastRow <= 1) {
      Logger.log("티커 데이터가 없습니다.");
      return [];
    }
    
    // Get data including header
    const tickerData = tickerSheet.getRange(1, 1, lastRow, lastCol).getValues();
    const result = [];
    
    // Skip header row
    for (let i = 1; i < tickerData.length; i++) {
      const row = tickerData[i];
      
      // Skip completely empty rows silently
      if (row.every(cell => cell === "")) {
        continue;
      }
      
      // Check if the row has the required data (name, ticker, source)
      if (!row[0] || !row[1] || !row[2]) {
        // Only log if there's some data in the row but missing required fields
        Logger.log(`잘못된 티커 데이터: 빈 값이 있습니다 (${row.join(', ')})`);
        continue;
      }
      
      // Validate the source is one of the supported data providers
      const validSources = ['google', 'yahoo', 'naver'];
      if (!validSources.includes(row[2].toLowerCase())) {
        Logger.log(`지원되지 않는 데이터 소스: "${row[2]}". 지원되는 소스: ${validSources.join(', ')}`);
        continue;
      }
      
      result.push({
        name: row[0],
        ticker: row[1],
        source: row[2].toLowerCase()
      });
    }
    
    Logger.log(`총 ${result.length}개의 티커를 처리합니다.`);
    return result;
  } catch (error) {
    Logger.log(`티커 목록 가져오기 오류: ${error.message}`);
    return [];
  }
}
