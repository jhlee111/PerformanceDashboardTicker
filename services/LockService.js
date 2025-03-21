/**
 * Performance Dashboard Ticker - Lock Service Module
 * 
 * This module provides functionality for managing script locks to prevent
 * multiple concurrent updates to the dashboard.
 */

// Lock-related constants
const LOCK_KEYS = {
  UPDATE_IN_PROGRESS: 'updateInProgress',
  LAST_UPDATE_TIME: 'lastUpdateTime',
  LAST_UPDATE_USER: 'lastUpdateUser'
};

// Lock timeout (5 minutes in milliseconds)
const LOCK_TIMEOUT_MS = 5 * 60 * 1000;

/**
 * Check if an update is already in progress
 * @return {boolean} True if update is in progress
 */
function isUpdateInProgress() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const updateInProgress = scriptProperties.getProperty(LOCK_KEYS.UPDATE_IN_PROGRESS);
    
    if (!updateInProgress) {
      return false;
    }
    
    // Check if lock is stale (more than 5 minutes old)
    const lastUpdateTime = parseInt(scriptProperties.getProperty(LOCK_KEYS.LAST_UPDATE_TIME) || '0');
    const currentTime = new Date().getTime();
    
    if (currentTime - lastUpdateTime > LOCK_TIMEOUT_MS) {
      // Lock is stale, clear it and return false
      Logger.log('스크립트 잠금이 오래되었습니다. 잠금을 해제합니다.');
      clearUpdateLock();
      return false;
    }
    
    // Lock is still valid
    return true;
  } catch (error) {
    Logger.log(`잠금 상태 확인 오류: ${error.message}`);
    return false;
  }
}

/**
 * Acquire a lock for dashboard update
 * @return {boolean} True if lock was acquired successfully
 */
function acquireUpdateLock() {
  try {
    // Check if update is already in progress
    if (isUpdateInProgress()) {
      const scriptProperties = PropertiesService.getScriptProperties();
      const lastUpdateUser = scriptProperties.getProperty(LOCK_KEYS.LAST_UPDATE_USER) || '다른 사용자';
      const lastUpdateTime = new Date(parseInt(scriptProperties.getProperty(LOCK_KEYS.LAST_UPDATE_TIME) || '0'));
      
      const formattedTime = Utilities.formatDate(lastUpdateTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      
      Logger.log(`이미 업데이트가 진행 중입니다. (${formattedTime}에 ${lastUpdateUser}가 시작)`);
      return false;
    }
    
    // Set update in progress
    const scriptProperties = PropertiesService.getScriptProperties();
    const currentUser = Session.getActiveUser().getEmail() || '알 수 없는 사용자';
    const currentTime = new Date().getTime();
    
    scriptProperties.setProperty(LOCK_KEYS.UPDATE_IN_PROGRESS, 'true');
    scriptProperties.setProperty(LOCK_KEYS.LAST_UPDATE_TIME, currentTime.toString());
    scriptProperties.setProperty(LOCK_KEYS.LAST_UPDATE_USER, currentUser);
    
    Logger.log(`업데이트 잠금을 획득했습니다. (사용자: ${currentUser})`);
    return true;
  } catch (error) {
    Logger.log(`잠금 획득 오류: ${error.message}`);
    return false;
  }
}

/**
 * Clear the update lock
 */
function clearUpdateLock() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.deleteProperty(LOCK_KEYS.UPDATE_IN_PROGRESS);
    Logger.log('업데이트 잠금을 해제했습니다.');
  } catch (error) {
    Logger.log(`잠금 해제 오류: ${error.message}`);
  }
}

/**
 * Force reset the lock, regardless of who has it
 * Used for emergency cases where the script might be stuck
 */
function forceResetLock() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // Get lock status before clearing
    const updateInProgress = scriptProperties.getProperty(LOCK_KEYS.UPDATE_IN_PROGRESS);
    const lastUpdateUser = scriptProperties.getProperty(LOCK_KEYS.LAST_UPDATE_USER) || '알 수 없는 사용자';
    const lastUpdateTime = parseInt(scriptProperties.getProperty(LOCK_KEYS.LAST_UPDATE_TIME) || '0');
    
    // Clear all lock-related properties
    scriptProperties.deleteProperty(LOCK_KEYS.UPDATE_IN_PROGRESS);
    scriptProperties.deleteProperty(LOCK_KEYS.LAST_UPDATE_TIME);
    scriptProperties.deleteProperty(LOCK_KEYS.LAST_UPDATE_USER);
    
    // Only log detailed message if there was actually a lock
    if (updateInProgress) {
      const formattedTime = Utilities.formatDate(new Date(lastUpdateTime), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      Logger.log(`잠금 강제 해제: ${formattedTime}에 ${lastUpdateUser}가 획득한 잠금을 강제로 해제했습니다.`);
    } else {
      Logger.log('잠금 강제 해제: 활성화된 잠금이 없었습니다.');
    }
    
    return {
      wasLocked: !!updateInProgress,
      previousUser: lastUpdateUser,
      previousTime: lastUpdateTime ? new Date(lastUpdateTime) : null
    };
  } catch (error) {
    Logger.log(`잠금 강제 해제 오류: ${error.message}`);
    return {
      wasLocked: false,
      error: error.message
    };
  }
}

/**
 * Get lock status information
 * @return {Object} Lock status information
 */
function getLockStatus() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const updateInProgress = scriptProperties.getProperty(LOCK_KEYS.UPDATE_IN_PROGRESS) === 'true';
    
    if (!updateInProgress) {
      return {
        locked: false,
        user: null,
        time: null,
        isStale: false
      };
    }
    
    const lastUpdateUser = scriptProperties.getProperty(LOCK_KEYS.LAST_UPDATE_USER) || '알 수 없는 사용자';
    const lastUpdateTime = parseInt(scriptProperties.getProperty(LOCK_KEYS.LAST_UPDATE_TIME) || '0');
    const currentTime = new Date().getTime();
    const isStale = (currentTime - lastUpdateTime > LOCK_TIMEOUT_MS);
    
    return {
      locked: true,
      user: lastUpdateUser,
      time: new Date(lastUpdateTime),
      isStale: isStale
    };
  } catch (error) {
    Logger.log(`잠금 상태 확인 오류: ${error.message}`);
    return {
      locked: false,
      error: error.message
    };
  }
} 