/**
 * Utilities for date calculations and handling
 */

/**
 * DateCalculator class for handling date calculations
 */
class DateCalculator {
  /**
   * @param {Date} referenceDate - The reference date for calculations
   */
  constructor(referenceDate) {
    this.referenceDate = referenceDate;
    Logger.log(`DateCalculator 초기화 - 참조일: ${this.formatDate(referenceDate)}`);
    
    // Store the current date for sanity checks
    this.currentDate = new Date();
    Logger.log(`현재 시스템 날짜: ${this.formatDate(this.currentDate)}`);
  }
  
  /**
   * Get the reference date
   * @return {Date} The reference date
   */
  getReferenceDate() {
    return this.referenceDate;
  }
  
  /**
   * Get a date that is specified number of weeks ago from the reference date
   * @param {number} weeks - Number of weeks ago
   * @return {Date} Date from weeks ago
   */
  getDateWeeksAgo(weeks) {
    Logger.log(`${weeks}주 전 날짜 계산 시작 - 참조일: ${this.formatDate(this.referenceDate)}`);
    const daysAgo = weeks * 7;
    const weekAgoDate = this.getLastTradingDate(this.getPastDate(daysAgo));
    Logger.log(`${weeks}주 전 날짜: ${this.formatDate(weekAgoDate)}`);
    return weekAgoDate;
  }
  
  /**
   * Get a date that is specified number of months ago from the reference date
   * @param {number} months - Number of months ago
   * @return {Date} Date from months ago
   */
  getDateMonthsAgo(months) {
    Logger.log(`${months}개월 전 날짜 계산 시작 - 참조일: ${this.formatDate(this.referenceDate)}`);
    
    // Clone the reference date to avoid modifying it
    const date = new Date(this.referenceDate);
    
    // Adjust the month
    date.setMonth(date.getMonth() - months);
    
    // If we've gone into the future relative to today, use today's date
    if (date > this.currentDate) {
      Logger.log(`경고: 계산된 과거 날짜(${this.formatDate(date)})가 현재 날짜(${this.formatDate(this.currentDate)})보다 미래입니다`);
      date.setTime(this.currentDate.getTime());
      
      // Again subtract the months from today instead
      date.setMonth(date.getMonth() - months);
    }
    
    // Get the last trading date
    const monthAgoDate = this.getLastTradingDate(date);
    Logger.log(`${months}개월 전 날짜: ${this.formatDate(monthAgoDate)}`);
    return monthAgoDate;
  }
  
  /**
   * Get the start date of the current year from the reference date
   * @return {Date} First day of the year
   */
  getYearStartDate() {
    Logger.log(`연초 날짜 계산 시작 - 참조 연도: ${this.referenceDate.getFullYear()}`);
    
    // If reference date is in future year (beyond current), use current year instead
    let year = this.referenceDate.getFullYear();
    const currentYear = this.currentDate.getFullYear();
    
    if (year > currentYear) {
      Logger.log(`경고: 참조 연도(${year})가 현재 연도(${currentYear})보다 미래입니다. 현재 연도를 사용합니다.`);
      year = currentYear;
    }
    
    const firstDayOfYear = new Date(year, 0, 1);
    Logger.log(`연초 기준일 (조정 전): ${this.formatDate(firstDayOfYear)}`);
    
    const ytdDate = this.getLastTradingDate(firstDayOfYear);
    Logger.log(`최종 연초 날짜: ${this.formatDate(ytdDate)}`);
    return ytdDate;
  }
  
  /**
   * Get a date in the past
   * @param {number} daysAgo - Number of days in the past
   * @return {Date} The past date
   */
  getPastDate(daysAgo) {
    const pastDate = new Date(this.referenceDate.getTime() - daysAgo * 24 * 60 * 60 * 1000);
    Logger.log(`${daysAgo}일 전 날짜: ${this.formatDate(pastDate)}`);
    
    // Check if the calculated past date is in the future (compared to today)
    // This could happen with future reference dates in testing environments
    if (pastDate > this.currentDate) {
      Logger.log(`경고: 계산된 과거 날짜(${this.formatDate(pastDate)})가 현재 날짜(${this.formatDate(this.currentDate)})보다 미래입니다`);
      Logger.log(`2025년 날짜를 사용하므로 과거 데이터를 가져올 때 현재 날짜를 사용합니다`);
      
      // Instead of using a future date, use a date relative to the current date
      const adjustedDate = new Date(this.currentDate.getTime() - daysAgo * 24 * 60 * 60 * 1000);
      Logger.log(`조정된 과거 날짜: ${this.formatDate(adjustedDate)}`);
      return adjustedDate;
    }
    
    return pastDate;
  }
  
  /**
   * Get the last trading date on or before the given date
   * @param {Date} date - The date to check
   * @return {Date} The last trading date
   */
  getLastTradingDate(date) {
    Logger.log(`거래일 조정 전 날짜: ${this.formatDate(date)}`);
    
    let attempts = 0;
    let originalDate = new Date(date);
    
    while (this.isWeekend(date) && attempts < 7) {
      date = new Date(date.getTime() - 24 * 60 * 60 * 1000); // Go back one day
      attempts++;
    }
    
    if (date.getTime() !== originalDate.getTime()) {
      Logger.log(`주말 조정: ${this.formatDate(originalDate)} → ${this.formatDate(date)}`);
    }
    
    return date;
  }
  
  /**
   * Check if a date is a weekend
   * @param {Date} date - The date to check
   * @return {boolean} True if the date is a weekend
   */
  isWeekend(date) {
    const day = date.getDay();
    const isWeekendDay = (day === 0 || day === 6); // Sunday = 0, Saturday = 6
    
    if (isWeekendDay) {
      Logger.log(`${this.formatDate(date)}은(는) 주말입니다 (${day === 0 ? '일요일' : '토요일'})`);
    }
    
    return isWeekendDay;
  }
  
  /**
   * Get the date one week ago, adjusted for trading days
   * @return {Date} The date one week ago
   */
  getWeekAgoDate() {
    Logger.log(`주간 날짜 계산 시작 - 참조일: ${this.formatDate(this.referenceDate)}`);
    const weekAgo = this.getLastTradingDate(this.getPastDate(CONFIG.DATE_RANGES.WEEKLY));
    Logger.log(`최종 주간 날짜: ${this.formatDate(weekAgo)}`);
    return weekAgo;
  }
  
  /**
   * Get the date one month ago, adjusted for trading days
   * @return {Date} The date one month ago
   */
  getMonthAgoDate() {
    Logger.log(`월간 날짜 계산 시작 - 참조일: ${this.formatDate(this.referenceDate)}`);
    const monthAgo = this.getLastTradingDate(this.getPastDate(CONFIG.DATE_RANGES.MONTHLY));
    Logger.log(`최종 월간 날짜: ${this.formatDate(monthAgo)}`);
    return monthAgo;
  }
  
  /**
   * Get the first trading date of the year
   * @return {Date} The first trading date of the year
   */
  getYtdDate() {
    Logger.log(`연초 날짜 계산 시작 - 참조 연도: ${this.referenceDate.getFullYear()}`);
    
    // If reference date is in future year (beyond current), use current year instead
    let year = this.referenceDate.getFullYear();
    const currentYear = this.currentDate.getFullYear();
    
    if (year > currentYear) {
      Logger.log(`경고: 참조 연도(${year})가 현재 연도(${currentYear})보다 미래입니다. 현재 연도를 사용합니다.`);
      year = currentYear;
    }
    
    const firstDayOfYear = new Date(year, 0, 1);
    Logger.log(`연초 기준일 (조정 전): ${this.formatDate(firstDayOfYear)}`);
    
    const ytdDate = this.getLastTradingDate(firstDayOfYear);
    Logger.log(`최종 연초 날짜: ${this.formatDate(ytdDate)}`);
    return ytdDate;
  }
  
  /**
   * Format a date as YYYY-MM-DD for logging
   * @param {Date} date - The date to format
   * @return {string} Formatted date string
   */
  formatDate(date) {
    if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
      return "유효하지 않은 날짜";
    }
    return Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  }
}

/**
 * Market Time Manager - Handles timezone-specific market status and availability
 */
class MarketTimeManager {
  constructor() {
    this.markets = {
      'kr': {
        name: '한국',
        timezone: 'Asia/Seoul',
        offset: 9, // GMT+9
        openTime: 9, // 9:00 AM
        closeTime: 15.5, // 3:30 PM (15:30)
        tradingDays: [1, 2, 3, 4, 5] // Monday to Friday
      },
      'us': {
        name: '미국',
        timezone: 'America/New_York',
        offset: -5, // EST (GMT-5) - may need seasonal adjustment for DST
        openTime: 9.5, // 9:30 AM
        closeTime: 16, // 4:00 PM
        tradingDays: [1, 2, 3, 4, 5] // Monday to Friday
      },
      'cn': {
        name: '중국',
        timezone: 'Asia/Shanghai',
        offset: 8, // GMT+8
        openTime: 9.5, // 9:30 AM
        closeTime: 15, // 3:00 PM
        tradingDays: [1, 2, 3, 4, 5] // Monday to Friday
      },
      'eu': {
        name: '유럽',
        timezone: 'Europe/London',
        offset: 0, // GMT (may need seasonal adjustment for DST)
        openTime: 8, // 8:00 AM (approx. for major European markets)
        closeTime: 16.5, // 4:30 PM (approx.)
        tradingDays: [1, 2, 3, 4, 5] // Monday to Friday
      }
    };
  }
  
  /**
   * Get the market region based on data source
   * @param {string} source - Data source (google, yahoo, naver)
   * @param {string} symbol - The ticker symbol
   * @returns {string} Market region code
   */
  getMarketRegion(source, symbol) {
    // Default to US market unless we can determine otherwise
    let region = 'us';
    
    if (source.toLowerCase() === 'naver') {
      // Naver is exclusively for Korean market
      return 'kr';
    }
    
    // Check symbol patterns
    if (symbol) {
      // Korean stocks typically have numeric symbols
      if (/^\d{6}$/.test(symbol)) {
        return 'kr';
      }
      
      // Chinese symbols often end with .SS or .SZ
      if (symbol.endsWith('.SS') || symbol.endsWith('.SZ') || symbol.endsWith('.HK')) {
        return 'cn';
      }
      
      // European symbols often include specific exchanges
      if (symbol.includes('.PA') || symbol.includes('.L') || symbol.includes('.F') || 
          symbol.includes('.MC') || symbol.includes('.AS') || symbol.includes('.MI')) {
        return 'eu';
      }
    }
    
    return region;
  }
  
  /**
   * Check if a market is open at a specific time
   * @param {string} marketRegion - Market region code (kr, us, cn, eu)
   * @param {Date} checkDate - Date to check
   * @returns {boolean} True if the market is open
   */
  isMarketOpen(marketRegion, checkDate) {
    const market = this.markets[marketRegion];
    if (!market) return false;
    
    // Convert check date to market timezone
    const marketDate = new Date(checkDate.getTime());
    const localOffset = -checkDate.getTimezoneOffset() / 60;
    const hourDiff = market.offset - localOffset;
    marketDate.setHours(marketDate.getHours() + hourDiff);
    
    // Check if it's a trading day
    const dayOfWeek = marketDate.getDay();
    if (!market.tradingDays.includes(dayOfWeek)) {
      return false;
    }
    
    // Check if it's during trading hours
    const marketHours = marketDate.getHours() + (marketDate.getMinutes() / 60);
    return marketHours >= market.openTime && marketHours < market.closeTime;
  }
  
  /**
   * Get the most recent trading day for a market
   * @param {string} marketRegion - Market region code (kr, us, cn, eu)
   * @param {Date} referenceDate - Reference date
   * @returns {Object} Date object with information about the most recent trading day
   */
  getMostRecentTradingDay(marketRegion, referenceDate) {
    const market = this.markets[marketRegion];
    if (!market) {
      return {
        date: referenceDate,
        isToday: true,
        isCurrent: true,
        message: '알 수 없는 시장'
      };
    }
    
    // Convert reference date to market timezone
    const marketDate = new Date(referenceDate.getTime());
    const localOffset = -referenceDate.getTimezoneOffset() / 60;
    const hourDiff = market.offset - localOffset;
    marketDate.setHours(marketDate.getHours() + hourDiff);
    
    const now = new Date();
    const marketNow = new Date(now.getTime());
    marketNow.setHours(marketNow.getHours() + hourDiff);
    
    // Check if today is a trading day
    const dayOfWeek = marketDate.getDay();
    const isTradingDay = market.tradingDays.includes(dayOfWeek);
    
    // Check market hours
    const marketHours = marketNow.getHours() + (marketNow.getMinutes() / 60);
    const isAfterClose = marketHours >= market.closeTime;
    const isBeforeOpen = marketHours < market.openTime;
    
    // Create result object
    let result = {
      date: new Date(marketDate),
      isToday: this.isSameDay(marketDate, marketNow),
      isCurrent: true,
      message: `${market.name} 시장 오늘 종가`
    };
    
    // Adjust based on market status
    if (!isTradingDay) {
      // Go back to the last trading day
      const daysToSubtract = this.getDaysToLastTradingDay(dayOfWeek, market.tradingDays);
      marketDate.setDate(marketDate.getDate() - daysToSubtract);
      
      result.date = new Date(marketDate);
      result.isToday = false;
      result.isCurrent = false;
      result.message = `${market.name} 시장 최근 거래일 (${daysToSubtract}일 전)`;
    } else if (isBeforeOpen) {
      // Market hasn't opened yet today, use previous close
      marketDate.setDate(marketDate.getDate() - 1);
      
      // Adjust if previous day is not a trading day
      let prevDayOfWeek = marketDate.getDay();
      if (!market.tradingDays.includes(prevDayOfWeek)) {
        const daysToSubtract = this.getDaysToLastTradingDay(prevDayOfWeek, market.tradingDays);
        marketDate.setDate(marketDate.getDate() - daysToSubtract);
      }
      
      result.date = new Date(marketDate);
      result.isToday = false;
      result.isCurrent = false;
      result.message = `${market.name} 시장 이전 거래일 (장 개장 전)`;
    } else if (!isAfterClose && result.isToday) {
      // Market is open but hasn't closed yet, use previous close
      marketDate.setDate(marketDate.getDate() - 1);
      
      // Adjust if previous day is not a trading day
      let prevDayOfWeek = marketDate.getDay();
      if (!market.tradingDays.includes(prevDayOfWeek)) {
        const daysToSubtract = this.getDaysToLastTradingDay(prevDayOfWeek, market.tradingDays);
        marketDate.setDate(marketDate.getDate() - daysToSubtract);
      }
      
      result.date = new Date(marketDate);
      result.isToday = false;
      result.isCurrent = false;
      result.message = `${market.name} 시장 이전 거래일 (장 중)`;
    }
    
    return result;
  }
  
  /**
   * Check if two dates are the same day
   * @param {Date} date1 - First date
   * @param {Date} date2 - Second date
   * @returns {boolean} True if same day
   * @private
   */
  isSameDay(date1, date2) {
    return date1.getFullYear() === date2.getFullYear() &&
           date1.getMonth() === date2.getMonth() &&
           date1.getDate() === date2.getDate();
  }
  
  /**
   * Calculate days to subtract to get to the last trading day
   * @param {number} currentDay - Current day of week (0-6)
   * @param {Array} tradingDays - Array of trading days
   * @returns {number} Days to subtract
   * @private
   */
  getDaysToLastTradingDay(currentDay, tradingDays) {
    let daysToSubtract = 1;
    let checkDay = (currentDay - 1 + 7) % 7;
    
    while (!tradingDays.includes(checkDay) && daysToSubtract < 10) {
      daysToSubtract++;
      checkDay = (checkDay - 1 + 7) % 7;
    }
    
    return daysToSubtract;
  }
  
  /**
   * Get formatted market status message
   * @param {string} marketRegion - Market region code
   * @param {Date} referenceDate - Reference date
   * @returns {string} Formatted status message
   */
  getMarketStatusMessage(marketRegion, referenceDate) {
    const tradingDay = this.getMostRecentTradingDay(marketRegion, referenceDate);
    return tradingDay.message;
  }
} 