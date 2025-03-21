/**
 * Performance Dashboard Ticker - UI Service Module
 * 
 * This module provides functionality for UI-related operations.
 */

/**
 * Show a loading indicator in the specified range
 * @param {Sheet} sheet - The sheet to update
 * @param {number} row - The row number
 * @param {number} column - The column number
 * @param {string} message - Loading message to display
 */
function showLoadingIndicator(sheet, row, column, message = "로딩 중...") {
  try {
    const cell = sheet.getRange(row, column);
    cell.setValue(message);
    // Use light yellow background for loading indicator
    cell.setBackground("#FFFDE7");
    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log(`로딩 표시기 업데이트 오류: ${error.message}`);
  }
}

/**
 * Hide loading indicator by clearing the cell
 * @param {Sheet} sheet - The sheet to update
 * @param {number} row - The row number
 * @param {number} column - The column number
 */
function hideLoadingIndicator(sheet, row, column) {
  try {
    const cell = sheet.getRange(row, column);
    cell.setValue("");
    cell.setBackground(null);
    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log(`로딩 표시기 제거 오류: ${error.message}`);
  }
}

/**
 * Format a cell based on its return value
 * @param {Sheet} sheet - The sheet to update
 * @param {number} row - The row number
 * @param {number} column - The column number
 * @param {number} returnValue - The return value
 */
function formatReturnCell(sheet, row, column, returnValue) {
  try {
    if (returnValue === null || isNaN(returnValue)) {
      return;
    }
    
    const cell = sheet.getRange(row, column);
    
    // Format to 2 decimal places with % sign
    cell.setNumberFormat("0.00%");
    
    // Set color based on value
    if (returnValue > 0) {
      cell.setFontColor("#388E3C"); // Green
    } else if (returnValue < 0) {
      cell.setFontColor("#D32F2F"); // Red
    } else {
      cell.setFontColor("#000000"); // Black for zero
    }
  } catch (error) {
    Logger.log(`셀 서식 오류: ${error.message}`);
  }
}

/**
 * Show an error message for a ticker
 * @param {Sheet} sheet - The sheet to update
 * @param {number} row - The row number
 * @param {number} infoColumn - The info column number
 * @param {string} errorMessage - The error message
 */
function showTickerError(sheet, row, infoColumn, errorMessage) {
  try {
    // Add error message to info column
    const infoCell = sheet.getRange(row, infoColumn);
    infoCell.setValue(`오류: ${errorMessage}`);
    infoCell.setFontColor("#D32F2F"); // Red color for errors
    
    // Set background color for visual indication
    infoCell.setBackground("#FFEBEE"); // Light red background
  } catch (error) {
    Logger.log(`오류 표시 실패: ${error.message}`);
  }
}

/**
 * Create and show the sidebar UI
 */
function showSidebar() {
  try {
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('성능 대시보드 도구')
      .setWidth(300);
    
    SpreadsheetApp.getUi().showSidebar(html);
    Logger.log('사이드바가 표시되었습니다.');
  } catch (error) {
    Logger.log(`사이드바 표시 오류: ${error.message}`);
    showErrorAlert('사이드바를 표시할 수 없습니다', error.message);
  }
}

/**
 * Show an alert with the given title and message
 * @param {string} title - Alert title
 * @param {string} message - Alert message
 */
function showAlert(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (error) {
    Logger.log(`알림 표시 오류: ${error.message}`);
  }
}

/**
 * Show an error alert with the given title and error message
 * @param {string} title - Alert title
 * @param {string} errorMessage - Error message
 */
function showErrorAlert(title, errorMessage) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      `⚠️ ${title}`,
      `오류가 발생했습니다: ${errorMessage}\n\n문제가 지속되면 관리자에게 문의하세요.`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    Logger.log(`오류 알림 표시 실패: ${error.message}`);
  }
}

/**
 * Show a confirmation dialog
 * @param {string} title - Dialog title
 * @param {string} message - Dialog message
 * @return {boolean} True if user clicked Yes, false otherwise
 */
function showConfirmation(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      title,
      message,
      ui.ButtonSet.YES_NO
    );
    
    return response === ui.Button.YES;
  } catch (error) {
    Logger.log(`확인 대화 상자 표시 오류: ${error.message}`);
    return false;
  }
}

/**
 * Create custom menu for dashboard
 */
function createCustomMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Create a menu with emojis for better visibility
    ui.createMenu('📊 대시보드')
      .addItem('🔄 대시보드 업데이트', 'updatePerformanceDashboard')
      .addItem('📅 기준일 설정', 'showReferenceDateSidebar')
      .addSeparator()
      .addSubMenu(ui.createMenu('⚙️ 관리')
        .addItem('📝 티커 관리', 'showSidebar')
        .addItem('🔬 진단 모드 켜기/끄기', 'toggleDiagnosticMode')
        .addItem('📋 디버그 보고서 생성', 'generateDebugReport')
        .addItem('🔓 잠금 강제 해제', 'resetLockWithConfirmation'))
      .addSeparator()
      .addSubMenu(ui.createMenu('❓ 도움말')
        .addItem('📚 사용 방법', 'showHelp')
        .addItem('🛠️ 문제 해결', 'showTroubleshooting')
        .addItem('ℹ️ 정보', 'showAbout'))
      .addToUi();
      
    Logger.log('성능 대시보드 메뉴가 생성되었습니다.');
  } catch (error) {
    Logger.log(`메뉴 생성 오류: ${error.message}`);
  }
}

/**
 * Show help information
 */
function showHelp() {
  const title = '📚 성능 대시보드 사용 방법';
  const message = 
    '이 대시보드는 다양한 주식 및 지수의 성과를 추적합니다.\n\n' +
    '기본 사용법:\n' +
    '1. "🔄 대시보드 업데이트" 메뉴를 클릭하여 최신 데이터로 업데이트합니다.\n' +
    '2. "📝 티커 관리" 메뉴를 사용하여 추적할 티커를 추가하거나 수정합니다.\n' +
    '3. "🔬 진단 모드"는 문제 해결을 위한 추가 로깅을 활성화합니다.\n\n' +
    '자세한 내용은 관리자에게 문의하세요.';
  
  showAlert(title, message);
}

/**
 * Show troubleshooting information
 */
function showTroubleshooting() {
  const title = '🛠️ 문제 해결';
  const message = 
    '일반적인 문제:\n\n' +
    '1. 업데이트가 진행 중일 때 "이미 업데이트가 진행 중입니다" 메시지가 표시됩니다.\n' +
    '   - 다른 사용자가 업데이트를 실행 중이거나, 이전 업데이트가 완료되지 않았습니다.\n' +
    '   - 5분 이상 기다린 후 다시 시도하거나, "🔓 잠금 강제 해제" 메뉴를 사용합니다.\n\n' +
    '2. 데이터가 표시되지 않거나 "N/A"로 표시됩니다.\n' +
    '   - 데이터 소스에서 정보를 가져올 수 없습니다.\n' +
    '   - 티커 정보를 확인하고 유효한지 확인하세요.\n\n' +
    '3. 오류 메시지가 표시됩니다.\n' +
    '   - "🔬 진단 모드"를 활성화하여 추가 로그를 확인하세요.\n' +
    '   - 스크립트 에디터(도구 > 스크립트 에디터)에서 로그를 확인할 수 있습니다.';
  
  showAlert(title, message);
}

/**
 * Show about information
 */
function showAbout() {
  const title = 'ℹ️ 성능 대시보드 정보';
  const message = 
    '성능 대시보드 티커\n' +
    '버전: 1.0.0\n\n' +
    '이 도구는 다양한 금융 상품의 성과를 추적하기 위해 개발되었습니다.\n' +
    '데이터는 Google Finance, Yahoo Finance, 네이버 금융에서 가져옵니다.\n\n' +
    '© 2023 All rights reserved\n';
  
  showAlert(title, message);
}

/**
 * Reset lock with confirmation
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

/**
 * Show the audit sheet
 */
function showAuditSheet() {
  try {
    // Initialize audit sheet if needed
    const auditSheet = initializeAuditSheet();
    
    if (!auditSheet) {
      showErrorAlert('감사 시트 오류', '감사 시트를 표시할 수 없습니다.');
      return;
    }
    
    // Activate the audit sheet
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(auditSheet);
    
    // Show information
    showAlert(
      '📋 현재 감사 로그', 
      '이 감사 로그에는 가장 최근 대시보드 업데이트에서의 데이터 소스 정보가 포함되어 있습니다.\n\n' +
      '이 시트를 통해 현재 대시보드에 표시된 데이터의 출처와 수집 방법을 확인할 수 있습니다.\n\n' +
      '대시보드가 업데이트될 때마다 이 감사 로그는 초기화되고 새로운 정보로 채워집니다.'
    );
  } catch (error) {
    Logger.log(`감사 시트 표시 오류: ${error.message}`);
    showErrorAlert('감사 시트 표시 오류', error.message);
  }
}

/**
 * Show sidebar with focus on reference date section
 */
function showReferenceDateSidebar() {
  try {
    // Create HTML output with a focus parameter
    const html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('기준일 설정')
      .setWidth(300);
    
    // Add a script to run when the sidebar loads to focus on the date section
    const script = 
      `<script>
        window.addEventListener('load', function() {
          // Focus on the date section
          const dateSection = document.getElementById('dateSection');
          if (dateSection) {
            // Scroll to the date section
            dateSection.scrollIntoView({ behavior: 'smooth' });
            
            // Highlight the section
            dateSection.style.boxShadow = '0 0 8px #4285f4';
            
            // Remove highlight after a delay
            setTimeout(function() {
              dateSection.style.boxShadow = 'none';
            }, 2000);
            
            // Focus on the date input
            const dateInput = document.getElementById('referenceDate');
            if (dateInput) {
              setTimeout(function() {
                dateInput.focus();
              }, 500);
            }
          }
        });
      </script>`;
    
    // Append the script to the HTML content
    html.append(script);
    
    // Show the sidebar
    SpreadsheetApp.getUi().showSidebar(html);
    
    // Log the action
    Logger.log('기준일 설정을 위한 사이드바가 표시되었습니다.');
  } catch (error) {
    Logger.log(`기준일 설정 사이드바 표시 오류: ${error.message}`);
    showErrorAlert('사이드바를 표시할 수 없습니다', error.message);
  }
} 