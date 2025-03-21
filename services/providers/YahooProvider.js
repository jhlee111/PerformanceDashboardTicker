/**
 * Yahoo Finance data provider
 */
class YahooFinanceProvider extends DataProvider {
  constructor(marketTimeManager) {
    super();
    this.name = "Yahoo Finance";
    this.marketTimeManager = marketTimeManager || new MarketTimeManager();
  }
  
  /**
   * Get the price for a symbol
   * @param {string} symbol - Ticker symbol
   * @param {Date} date - Date for price data (used for historical data)
   * @return {number} The price
   */
  getPrice(symbol, date) {
    if (!date || this.isToday(date)) {
      return this.getCurrentPrice(symbol);
    } else {
      return this.getHistoricalPrice(symbol, date);
    }
  }
  
  /**
   * Check if a date is today
   * @param {Date} date - The date to check
   * @return {boolean} True if the date is today
   */
  isToday(date) {
    const today = new Date();
    return date.getFullYear() === today.getFullYear() &&
           date.getMonth() === today.getMonth() &&
           date.getDate() === today.getDate();
  }
  
  /**
   * Encode a symbol for use in Yahoo Finance URLs
   * @param {string} symbol - The symbol to encode
   * @return {string} The encoded symbol
   */
  encodeSymbol(symbol) {
    // Replace special characters that need encoding in URLs
    return encodeURIComponent(symbol);
  }
  
  /**
   * Get the current price for a symbol
   * @param {string} symbol - Ticker symbol
   * @return {number} The current price
   */
  getCurrentPrice(symbol) {
    try {
      const encodedSymbol = this.encodeSymbol(symbol);
      const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodedSymbol}?interval=1d&range=1d`;
      const options = {
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36'
        }
      };
      
      const response = UrlFetchApp.fetch(url, options).getContentText();
      const json = JSON.parse(response);
      
      // First check if there's an error in the response
      if (json.chart && json.chart.error) {
        Logger.log(`Yahoo Finance API 오류: ${symbol} - ${json.chart.error.description}`);
        return 0;
      }
      
      if (json.chart && json.chart.result && json.chart.result.length > 0) {
        // First try to get regularMarketPrice from meta
        if (json.chart.result[0].meta && json.chart.result[0].meta.regularMarketPrice !== undefined) {
          return json.chart.result[0].meta.regularMarketPrice;
        }
        
        // If not available in meta, try to get the current price from quotes
        if (json.chart.result[0].indicators && 
            json.chart.result[0].indicators.quote && 
            json.chart.result[0].indicators.quote.length > 0) {
          
          const quote = json.chart.result[0].indicators.quote[0];
          
          if (quote.close && quote.close.length > 0) {
            // Get the last valid close price
            for (let i = quote.close.length - 1; i >= 0; i--) {
              if (quote.close[i] !== null && !isNaN(quote.close[i])) {
                return quote.close[i];
              }
            }
          }
        }
        
        // Try to get from additional quote data if available
        if (json.chart.result[0].meta && json.chart.result[0].meta.chartPreviousClose) {
          return json.chart.result[0].meta.chartPreviousClose;
        }
      }
      
      Logger.log(`Yahoo Finance API에서 현재 가격 정보가 없음: ${symbol}`);
    } catch (e) {
      Logger.log(`Yahoo Finance 현재 가격 가져오기 실패: ${symbol} - ${e.message}`);
    }
    
    return 0;
  }
  
  /**
   * Get a historical price for a symbol
   * @param {string} symbol - Ticker symbol
   * @param {Date} date - The date for the price
   * @return {number} The historical price
   */
  getHistoricalPrice(symbol, date) {
    try {
      // Format date as YYYY-MM-DD
      const formattedDate = Utilities.formatDate(date, "GMT", "yyyy-MM-dd");
      
      // Convert date to UNIX timestamp (seconds)
      const period1 = Math.floor(date.getTime() / 1000);
      
      // Add one day to get data for the specific date
      const nextDay = new Date(date);
      nextDay.setDate(nextDay.getDate() + 1);
      const period2 = Math.floor(nextDay.getTime() / 1000);
      
      const encodedSymbol = this.encodeSymbol(symbol);
      const url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodedSymbol}?interval=1d&period1=${period1}&period2=${period2}`;
      const options = {
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36'
        }
      };
      
      const response = UrlFetchApp.fetch(url, options).getContentText();
      const json = JSON.parse(response);
      
      // Check for API errors
      if (json.chart && json.chart.error) {
        Logger.log(`Yahoo Finance API 오류: ${symbol} - ${json.chart.error.description}`);
        return 0;
      }
      
      if (json.chart && json.chart.result && json.chart.result.length > 0) {
        // Try to get the adjusted close price
        if (json.chart.result[0].indicators && 
            json.chart.result[0].indicators.adjclose && 
            json.chart.result[0].indicators.adjclose.length > 0 &&
            json.chart.result[0].indicators.adjclose[0].adjclose &&
            json.chart.result[0].indicators.adjclose[0].adjclose.length > 0) {
            
          for (let i = 0; i < json.chart.result[0].indicators.adjclose[0].adjclose.length; i++) {
            if (json.chart.result[0].indicators.adjclose[0].adjclose[i] !== null && 
                !isNaN(json.chart.result[0].indicators.adjclose[0].adjclose[i])) {
              return json.chart.result[0].indicators.adjclose[0].adjclose[i];
            }
          }
        }
        
        // If adjusted close is not available, try regular close
        if (json.chart.result[0].indicators && 
            json.chart.result[0].indicators.quote && 
            json.chart.result[0].indicators.quote.length > 0) {
          
          const quote = json.chart.result[0].indicators.quote[0];
          
          if (quote.close && quote.close.length > 0) {
            // Return the first valid close price
            for (let i = 0; i < quote.close.length; i++) {
              if (quote.close[i] !== null && !isNaN(quote.close[i])) {
                return quote.close[i];
              }
            }
          }
        }
      }
      
      Logger.log(`Yahoo Finance API에서 날짜(${formattedDate})에 대한 가격 정보가 없음: ${symbol}`);
    } catch (e) {
      Logger.log(`Yahoo Finance 과거 가격 가져오기 실패: ${symbol} - ${e.message}`);
    }
    
    return 0;
  }
  
  /**
   * Get the highest price for a symbol
   * @param {string} symbol - Ticker symbol
   * @return {number} The highest price in the last 52 weeks
   */
  getHighPrice(symbol) {
    try {
      const encodedSymbol = this.encodeSymbol(symbol);
      
      // First try the quote endpoint which is more reliable for 52-week high
      const quoteUrl = `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${encodedSymbol}`;
      const options = {
        muteHttpExceptions: true,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36'
        }
      };
      
      const quoteResponse = UrlFetchApp.fetch(quoteUrl, options).getContentText();
      const quoteJson = JSON.parse(quoteResponse);
      
      if (quoteJson.quoteResponse && 
          quoteJson.quoteResponse.result && 
          quoteJson.quoteResponse.result.length > 0 &&
          quoteJson.quoteResponse.result[0].fiftyTwoWeekHigh) {
        return quoteJson.quoteResponse.result[0].fiftyTwoWeekHigh;
      }
      
      // Fallback to chart API if quote doesn't have the data
      const chartUrl = `https://query1.finance.yahoo.com/v8/finance/chart/${encodedSymbol}?interval=1d&range=1y`;
      
      const chartResponse = UrlFetchApp.fetch(chartUrl, options).getContentText();
      const chartJson = JSON.parse(chartResponse);
      
      if (chartJson.chart && chartJson.chart.error) {
        Logger.log(`Yahoo Finance API 오류: ${symbol} - ${chartJson.chart.error.description}`);
        return 0;
      }
      
      if (chartJson.chart && chartJson.chart.result && chartJson.chart.result.length > 0) {
        // Try to get the 52-week high from meta data
        if (chartJson.chart.result[0].meta && chartJson.chart.result[0].meta.fiftyTwoWeekHigh) {
          return chartJson.chart.result[0].meta.fiftyTwoWeekHigh;
        }
        
        // If not available in meta, calculate from historical data
        if (chartJson.chart.result[0].indicators && 
            chartJson.chart.result[0].indicators.quote && 
            chartJson.chart.result[0].indicators.quote.length > 0) {
          
          const quote = chartJson.chart.result[0].indicators.quote[0];
          
          if (quote.high && quote.high.length > 0) {
            // Find the highest value in the array
            let highestPrice = 0;
            for (let i = 0; i < quote.high.length; i++) {
              if (quote.high[i] !== null && !isNaN(quote.high[i]) && quote.high[i] > highestPrice) {
                highestPrice = quote.high[i];
              }
            }
            if (highestPrice > 0) {
              return highestPrice;
            }
          }
        }
      }
      
      Logger.log(`Yahoo Finance API에서 52주 최고가 정보가 없음: ${symbol}`);
    } catch (e) {
      Logger.log(`Yahoo Finance 52주 최고가 가져오기 실패: ${symbol} - ${e.message}`);
    }
    
    return 0;
  }
} 