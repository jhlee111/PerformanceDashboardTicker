/**
 * Performance Dashboard Ticker - Lock Service Module
 * 
 * This module provides functionality for script locking to prevent race conditions.
 */

/**
 * Get a script lock for preventing race conditions
 * @return {Lock|null} The script lock or null if failed
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
 * Force reset the system lock and update status
 * Use this when the script gets stuck
 */
function forceResetSystem() {
  try {
    clearUpdateStatus();
    
    // Try to release any existing script locks
    try {
      const lock = LockService.getScriptLock();
      if (lock) {
        lock.releaseLock();
      }
    } catch (lockError) {
      // Ignore lock release errors
    }
    
    SpreadsheetApp.getUi().alert(
      '시스템 초기화 완료', 
      '스크립트 시스템이 초기화되었습니다.\n이제 대시보드 업데이트를 다시 실행할 수 있습니다.', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    Logger.log('스크립트 시스템이 강제로 초기화되었습니다.');
    return true;
  } catch (error) {
    Logger.log(`시스템 초기화 실패: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      '시스템 초기화 실패', 
      `스크립트 시스템 초기화 중 오류가 발생했습니다: ${error.message}`, 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
} 