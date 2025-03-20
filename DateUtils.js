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