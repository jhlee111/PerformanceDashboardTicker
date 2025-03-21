/**
 * Performance Dashboard Ticker - Sidebar UI Module
 * 
 * This module contains functions for creating and managing the UI sidebar.
 */

/**
 * Show the sidebar for user interaction
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('성과 대시보드 컨트롤')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Update the dashboard when called from the sidebar
 * Simple wrapper for calling from the sidebar UI
 */
function updateDashboardFromSidebar() {
  try {
    updatePerformanceDashboard();
    return { 
      success: true, 
      message: "대시보드가 성공적으로 업데이트되었습니다."
    };
  } catch (error) {
    Logger.log(`사이드바에서 대시보드 업데이트 오류: ${error.message}`);
    return { 
      success: false, 
      message: `대시보드 업데이트 중 오류가 발생했습니다: ${error.message}`
    };
  }
}

/**
 * Set reference date from sidebar
 * @param {string} dateString - Date string in YYYY-MM-DD format
 * @return {Object} Result object with success status and message
 */
function setReferenceDateFromSidebar(dateString) {
  try {
    const success = setReferenceDate(dateString);
    
    if (success) {
      return { 
        success: true, 
        message: "참조일이 성공적으로 설정되었습니다."
      };
    } else {
      return { 
        success: false, 
        message: "참조일을 설정할 수 없습니다. 날짜 형식을 확인해주세요."
      };
    }
  } catch (error) {
    Logger.log(`사이드바에서 참조일 설정 오류: ${error.message}`);
    return { 
      success: false, 
      message: `참조일 설정 중 오류가 발생했습니다: ${error.message}`
    };
  }
}

/**
 * Get current reference date for sidebar
 * @return {Object} Result object with date string
 */
function getCurrentReferenceDateForSidebar() {
  try {
    const refDate = getReferenceDate();
    const dateString = Utilities.formatDate(refDate, "GMT+9", "yyyy-MM-dd");
    
    return { 
      success: true, 
      date: dateString
    };
  } catch (error) {
    Logger.log(`사이드바를 위한 참조일 가져오기 오류: ${error.message}`);
    return { 
      success: false, 
      message: `참조일을 가져오는 중 오류가 발생했습니다: ${error.message}`
    };
  }
}

/**
 * Get current progress data for sidebar
 * @return {Object} Progress data
 */
function getProgressForSidebar() {
  return getProgress();
} 