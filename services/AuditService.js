/**
 * Performance Dashboard Ticker - Audit Service Module
 * 
 * This module provides functionality for tracking and auditing data fetch operations.
 */

/**
 * Initialize the audit sheet
 * @return {Sheet} The audit sheet
 */
function initializeAuditSheet() {
  try {
    Logger.log('감사 시트 초기화 중...');
    
    // Get or create audit sheet
    let auditSheet = SS.getSheetByName(CONFIG.SHEETS.AUDIT);
    if (!auditSheet) {
      auditSheet = SS.insertSheet(CONFIG.SHEETS.AUDIT);
      Logger.log('감사 시트를 생성했습니다.');
    }
    
    // Define headers
    const headers = [
      '타임스탬프', '티커 심볼', '티커 이름', '데이터 소스', 
      '조회 날짜 (시장별)', '참조 날짜', '조회 방법', '현재가', 
      '주간 가격', '월간 가격', 'YTD 가격', '52주 최고가', 
      '주간 변화', '월간 변화', 'YTD 변화', '최고가 대비',
      '추가 정보'
    ];
    
    // Check if headers already exist
    const existingHeaders = auditSheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const headersMatch = existingHeaders.every((header, index) => header === headers[index]);
    
    if (!headersMatch) {
      // Set headers
      const headerRange = auditSheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      
      // Format headers
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#E3F2FD'); // Light blue background
      
      // Add notes to headers for clarity
      auditSheet.getRange(1, 5).setNote('조회 날짜는 해당 시장의 시간대를 반영한 가장 최근 거래일입니다. 시장별로 다를 수 있습니다.');
      auditSheet.getRange(1, 6).setNote('참조 날짜는 사용자가 설정한 기준일입니다. 모든 시장에 동일하게 적용됩니다.');
      
      // Freeze the header row
      auditSheet.setFrozenRows(1);
      
      // Auto-size columns
      for (let i = 1; i <= headers.length; i++) {
        auditSheet.autoResizeColumn(i);
      }
      
      Logger.log('감사 시트 헤더를 설정했습니다.');
    }
    
    return auditSheet;
  } catch (error) {
    Logger.log(`감사 시트 초기화 오류: ${error.message}`);
    return null;
  }
}

/**
 * Log data fetching details to the audit sheet
 * @param {Object} data - Data to log
 */
function logAuditData(data) {
  try {
    if (!data) {
      Logger.log('감사 로그를 위한 데이터가 제공되지 않았습니다.');
      return;
    }
    
    // Initialize audit sheet
    const auditSheet = initializeAuditSheet();
    if (!auditSheet) {
      Logger.log('감사 시트를 초기화할 수 없어 로그를 작성할 수 없습니다.');
      return;
    }
    
    // Find the next available row
    const lastRow = auditSheet.getLastRow();
    const nextRow = lastRow + 1;
    
    // Format returns for readability (if they exist)
    const formatReturnForAudit = (returnValue) => {
      if (returnValue === null || returnValue === undefined || isNaN(returnValue)) {
        return 'N/A';
      }
      const sign = returnValue > 0 ? '+' : '';
      return `${sign}${returnValue.toFixed(2)}%`;
    };
    
    // Prepare row data
    const rowData = [
      new Date(), // Timestamp
      data.symbol, 
      data.name || '',
      data.source || '',
      data.fetchDate ? Utilities.formatDate(data.fetchDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') + 
        ` (${getMarketRegionDisplay(data.source, data.symbol)})` : '',
      data.referenceDate ? Utilities.formatDate(data.referenceDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      data.method || '직접 조회',
      data.prices?.current || 'N/A',
      data.prices?.weekly || 'N/A',
      data.prices?.monthly || 'N/A',
      data.prices?.ytd || 'N/A',
      data.prices?.high || 'N/A',
      formatReturnForAudit(data.returns?.weekly),
      formatReturnForAudit(data.returns?.monthly),
      formatReturnForAudit(data.returns?.ytd),
      formatReturnForAudit(data.returns?.high),
      data.notes || ''
    ];
    
    // Write data to sheet
    auditSheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
    
    // Format the timestamp
    auditSheet.getRange(nextRow, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
    
    // Format the price columns
    for (let i = 8; i <= 12; i++) {
      auditSheet.getRange(nextRow, i).setNumberFormat('#,##0.00');
    }
    
    Logger.log(`감사 데이터가 행 ${nextRow}에 기록되었습니다: ${data.symbol} (${data.source})`);
  } catch (error) {
    Logger.log(`감사 데이터 로깅 오류: ${error.message}`);
  }
}

/**
 * Add market-specific notes based on the data
 * @param {string} symbol - The ticker symbol
 * @param {string} source - The data source
 * @param {Object} prices - Price data
 * @param {MarketTimeManager} marketTimeManager - Market time manager
 * @return {string} Notes about the data
 */
function generateAuditNotes(symbol, source, prices, marketTimeManager) {
  let notes = [];
  
  try {
    // Get market region
    const marketRegion = marketTimeManager.getMarketRegion(source, symbol);
    const regionDisplay = getMarketRegionName(marketRegion);
    
    // Get most recent trading day information
    const currentDate = new Date();
    const tradingDayInfo = marketTimeManager.getMostRecentTradingDay(marketRegion, currentDate);
    const isMarketOpen = marketTimeManager.isMarketOpen(marketRegion, currentDate);
    
    // Add market region info
    notes.push(`시장: ${regionDisplay}`);
    
    // Check for potential market closure or data availability issues
    if (!isMarketOpen) {
      const tradingDateStr = Utilities.formatDate(tradingDayInfo.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      notes.push(`${regionDisplay} 시장은 현재 닫혀 있습니다. ${tradingDateStr} 거래일 데이터를 사용합니다.`);
    }
    
    // Check for missing historical data
    if (!prices.weekly || isNaN(prices.weekly)) {
      notes.push('주간 가격 데이터를 가져올 수 없습니다.');
    }
    
    if (!prices.monthly || isNaN(prices.monthly)) {
      notes.push('월간 가격 데이터를 가져올 수 없습니다.');
    }
    
    if (!prices.ytd || isNaN(prices.ytd)) {
      notes.push('연초 가격 데이터를 가져올 수 없습니다.');
    }
    
    // Add source-specific notes
    if (source === 'google') {
      notes.push('Google Finance API를 통해 가져온 데이터입니다.');
    } else if (source === 'yahoo') {
      notes.push('Yahoo Finance API를 통해 가져온 데이터입니다.');
    } else if (source === 'naver') {
      notes.push('네이버 금융에서 파싱한 데이터입니다.');
      
      // Add note for KOSDAQ special handling
      if (symbol === 'KOSDAQ') {
        notes.push('KOSDAQ 지수는 기록 데이터 부재로 인해 추정 가격이 사용되었을 수 있습니다.');
      }
    }
  } catch (error) {
    notes.push(`감사 노트 생성 중 오류: ${error.message}`);
  }
  
  return notes.join(' ');
}

/**
 * Get readable market region name
 * @param {string} marketRegion - Market region code
 * @return {string} Human readable market region name
 */
function getMarketRegionName(marketRegion) {
  switch(marketRegion) {
    case 'us':
      return '미국';
    case 'kr':
      return '한국';
    case 'cn':
      return '중국';
    case 'eu':
      return '유럽';
    default:
      return '기타';
  }
}

/**
 * Clear all audit data from the audit sheet
 * @return {Sheet} The cleared audit sheet
 */
function clearAuditData() {
  try {
    Logger.log('감사 데이터 초기화 중...');
    
    // Get or create audit sheet
    let auditSheet = SS.getSheetByName(CONFIG.SHEETS.AUDIT);
    
    if (auditSheet) {
      // Instead of deleting the sheet, just clear its contents
      // Get the number of rows with data
      const lastRow = auditSheet.getLastRow();
      const lastCol = auditSheet.getLastColumn();
      
      // If there's data beyond the header row, clear it
      if (lastRow > 1) {
        const dataRange = auditSheet.getRange(2, 1, lastRow - 1, lastCol);
        dataRange.clearContent();
        
        Logger.log(`기존 감사 시트의 데이터를 초기화했습니다 (${lastRow - 1}행 제거).`);
      } else {
        Logger.log('감사 시트에 지울 데이터가 없습니다.');
      }
    } else {
      // Create a new audit sheet if it doesn't exist
      auditSheet = SS.insertSheet(CONFIG.SHEETS.AUDIT);
      Logger.log('새 감사 시트가 생성되었습니다.');
    }
    
    // Define headers
    const headers = [
      '타임스탬프', '티커 심볼', '티커 이름', '데이터 소스', 
      '조회 날짜 (시장별)', '참조 날짜', '조회 방법', '현재가', 
      '주간 가격', '월간 가격', 'YTD 가격', '52주 최고가', 
      '주간 변화', '월간 변화', 'YTD 변화', '최고가 대비',
      '추가 정보'
    ];
    
    // Check if headers already exist and match expected headers
    const firstRow = auditSheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const headersMatch = firstRow.every((header, index) => header === headers[index]);
    
    // Only update headers if they don't match
    if (!headersMatch) {
      // Set headers
      const headerRange = auditSheet.getRange(1, 1, 1, headers.length);
      headerRange.setValues([headers]);
      
      // Format headers
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#E3F2FD'); // Light blue background
      
      // Add notes to headers for clarity
      auditSheet.getRange(1, 5).setNote('조회 날짜는 해당 시장의 시간대를 반영한 가장 최근 거래일입니다. 시장별로 다를 수 있습니다.');
      auditSheet.getRange(1, 6).setNote('참조 날짜는 사용자가 설정한 기준일입니다. 모든 시장에 동일하게 적용됩니다.');
      
      // Freeze the header row
      auditSheet.setFrozenRows(1);
      
      Logger.log('감사 시트 헤더를 업데이트했습니다.');
    }
    
    // Auto-size columns - always do this to ensure proper display
    for (let i = 1; i <= headers.length; i++) {
      auditSheet.autoResizeColumn(i);
    }
    
    Logger.log('감사 시트가 초기화되었습니다.');
    
    return auditSheet;
  } catch (error) {
    Logger.log(`감사 데이터 초기화 오류: ${error.message}`);
    return null;
  }
}

/**
 * Get market region display name for the audit log
 * @param {string} source - Data source
 * @param {string} symbol - Ticker symbol
 * @return {string} Market region display name
 */
function getMarketRegionDisplay(source, symbol) {
  const marketTimeManager = new MarketTimeManager();
  const region = marketTimeManager.getMarketRegion(source, symbol);
  
  switch(region) {
    case 'us':
      return 'US/EDT';
    case 'kr':
      return 'KR/KST';
    case 'cn':
      return 'CN/CST';
    case 'eu':
      return 'EU/CET';
    default:
      return '';
  }
} 