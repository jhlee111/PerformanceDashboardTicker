/**
 * Performance Dashboard Ticker - Main Module
 * 
 * Entry point and initialization for the dashboard application.
 */

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
 * Create the menu when the spreadsheet is opened
 */
function onOpen() {
  // Create custom menu
  createCustomMenu();
  
  Logger.log('스프레드시트가 열렸습니다.');
}

/**
 * Initialize global settings when the spreadsheet is opened
 * This is automatically called when the spreadsheet is opened
 */
function initialize() {
  // Initialize progress tracking
  PROGRESS_DATA = {
    percent: 0,
    message: "초기화 중...",
    type: "progress",
    tickers: []
  };
  
  // Make sure the global spreadsheet instance is always up to date
  SS = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log('Performance Dashboard initialized successfully');
}

/**
 * Initialize the application and set up required resources
 * @return {boolean} True if initialization is successful
 */
function initializeApp() {
  try {
    // Initialize progress tracking
    PROGRESS_DATA = {
      percent: 0,
      message: "초기화 중...",
      type: "progress",
      tickers: []
    };
    
    // Ensure all required sheets exist
    const dashboardSheet = getDashboardSheet();
    if (!dashboardSheet) {
      throw new Error("대시보드 시트를 찾을 수 없습니다.");
    }
    
    // Ensure temp calculation sheet exists
    let tempSheet = SS.getSheetByName(CONFIG.SHEETS.TEMP);
    if (!tempSheet) {
      // Create temp sheet if it doesn't exist
      tempSheet = SS.insertSheet(CONFIG.SHEETS.TEMP);
      // Hide the temp sheet
      tempSheet.hideSheet();
    }
    
    // Other initialization tasks can be added here
    
    return true;
  } catch (error) {
    Logger.log(`애플리케이션 초기화 오류: ${error.message}`);
    showErrorAlert('애플리케이션 초기화 오류', error.message);
    return false;
  }
}

/**
 * Enable diagnostic mode (on)
 */
function enableDiagnosticModeOn() {
  enableDiagnosticMode(true);
}

/**
 * Disable diagnostic mode (off)
 */
function enableDiagnosticModeOff() {
  enableDiagnosticMode(false);
}

/**
 * Toggle diagnostic mode
 */
function toggleDiagnosticMode() {
  const currentStatus = isDiagnosticModeEnabled();
  enableDiagnosticMode(!currentStatus, true); // Show alert when toggling via menu
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
 * Get the current timestamp as a formatted string
 * @return {string} Formatted timestamp (e.g., "2025.03.21 17:30:45")
 */
function getCurrentTimestamp() {
  const now = new Date();
  return Utilities.formatDate(now, "GMT+9", "yyyy.MM.dd HH:mm:ss");
}

/**
 * Get reference date (usually today)
 * @return {Date} Reference date
 */
function getReferenceDate() {
  try {
    // Use current date as reference
    return new Date();
  } catch (error) {
    Logger.log(`기준일 가져오기 오류: ${error.message}`);
    return new Date();
  }
}

/**
 * Get current reference date for display in UI
 * @return {string} Formatted reference date (YYYY-MM-DD)
 */
function getCurrentReferenceDate() {
  try {
    const date = getReferenceDate();
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (error) {
    Logger.log(`참조일 가져오기 오류: ${error.message}`);
    return "오류";
  }
}

/**
 * Set reference date from the sidebar
 * @param {string} dateString - Date string in YYYY-MM-DD format
 * @return {Object} Result object with success flag and messages
 */
function setReferenceDate(dateString) {
  try {
    // Validate date format
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateString)) {
      return {
        success: false,
        message: '유효하지 않은 날짜 형식입니다. YYYY-MM-DD 형식을 사용하세요.'
      };
    }
    
    // Parse date parts
    const [year, month, day] = dateString.split('-').map(part => parseInt(part, 10));
    
    // Check valid date values
    if (year < 2000 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) {
      return {
        success: false,
        message: '유효하지 않은 날짜입니다. 유효한 날짜를 입력하세요.'
      };
    }
    
    // Create new date object
    const newDate = new Date(year, month - 1, day); // Month is 0-indexed in JavaScript
    
    // Store in script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('REFERENCE_DATE', newDate.toISOString());
    
    // Format for consistent display
    const formattedDate = Utilities.formatDate(newDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    Logger.log(`참조일이 설정되었습니다: ${formattedDate}`);
    
    return {
      success: true,
      date: formattedDate,
      message: `참조일이 ${formattedDate}로 설정되었습니다.`
    };
  } catch (error) {
    Logger.log(`참조일 설정 오류: ${error.message}`);
    
    return {
      success: false,
      message: `참조일 설정 중 오류가 발생했습니다: ${error.message}`
    };
  }
}

/**
 * Reset reference date to today
 * @return {Object} Result with success flag and date
 */
function resetReferenceDate() {
  try {
    // Clear stored date to default to today
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty('REFERENCE_DATE');
    
    // Get the current date for confirmation
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    Logger.log(`참조일이 오늘(${formattedDate})로 재설정되었습니다.`);
    
    return {
      success: true,
      date: formattedDate,
      message: '참조일이 오늘로 재설정되었습니다.'
    };
  } catch (error) {
    Logger.log(`참조일 재설정 오류: ${error.message}`);
    
    return {
      success: false,
      message: `참조일 재설정 중 오류가 발생했습니다: ${error.message}`
    };
  }
}

/**
 * Reset lock with confirmation - convenient wrapper function
 */
function resetLockWithConfirmation() {
  try {
    const confirmReset = showConfirmation(
      '🔓 잠금 강제 해제',
      '정말로 잠금을 강제로 해제하시겠습니까?\n\n' +
      '이 작업은 현재 진행 중인 업데이트를 중단시킬 수 있습니다.\n' +
      '다른 사용자가 업데이트를 실행 중인 경우에만 사용하세요.'
    );
    
    if (!confirmReset) {
      return;
    }
    
    const lockStatus = forceResetLock();
    
    if (lockStatus.wasLocked) {
      const formattedTime = Utilities.formatDate(
        lockStatus.previousTime, 
        Session.getScriptTimeZone(), 
        'yyyy-MM-dd HH:mm:ss'
      );
      
      showAlert(
        '🔓 잠금 해제됨',
        `${formattedTime}에 ${lockStatus.previousUser}가 획득한 잠금이 해제되었습니다.\n\n` +
        '이제 업데이트를 실행할 수 있습니다.'
      );
    } else {
      showAlert(
        '🔓 잠금 상태',
        '활성화된 잠금이 없습니다. 업데이트를 실행할 수 있습니다.'
      );
    }
  } catch (error) {
    Logger.log(`잠금 강제 해제 오류: ${error.message}`);
    showErrorAlert('잠금 강제 해제 실패', error.message);
  }
} 