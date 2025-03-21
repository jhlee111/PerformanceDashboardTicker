/**
 * PlayWright-based integration test for KOSDAQ historical data fetching
 * This test uses PlayWright to directly access Naver's mobile site
 * 
 * To run this:
 * 1. Install PlayWright CLI: npm install -g playwright
 * 2. Run: node __tests__/integration/naverKosdaqPlaywright.test.js
 */

const { chromium } = require('playwright');

// Configuration
const symbolCode = 'KOSDAQ';

/**
 * Run PlayWright test to validate KOSDAQ data
 */
async function runPlaywrightTest() {
  console.log('Starting PlayWright test for KOSDAQ historical data...');
  
  // Launch browser
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({
    userAgent: 'Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1',
    viewport: { width: 375, height: 812 },
    deviceScaleFactor: 3,
    isMobile: true,
    hasTouch: true
  });
  
  try {
    // Open a new page
    const page = await context.newPage();
    console.log('Opening Naver mobile stock page...');
    
    // Navigate to KOSDAQ page
    await page.goto(`https://m.stock.naver.com/domestic/index/${symbolCode}/total`, { waitUntil: 'networkidle' });
    console.log('Page loaded successfully');
    
    // Wait for price to appear
    await page.waitForSelector('div[class*="StockPrice"]', { timeout: 10000 });
    
    // Get current price
    const currentPrice = await page.evaluate(() => {
      const priceElement = document.querySelector('div[class*="StockPrice"] strong');
      return priceElement ? priceElement.textContent.replace(/,/g, '') : null;
    });
    
    console.log(`Current KOSDAQ price: ${currentPrice}`);
    
    // Go to chart page
    console.log('Navigating to chart page...');
    await page.click('a[href*="/chart"]');
    await page.waitForSelector('div[class*="ChartHeader"]', { timeout: 10000 });
    
    // Wait for chart data to load
    await page.waitForTimeout(3000);
    
    // Collect historical data through network requests
    console.log('Collecting API responses...');
    let historicalData = [];
    
    page.on('response', async (response) => {
      const url = response.url();
      if (url.includes('/api/index/KOSDAQ/price') && response.status() === 200) {
        try {
          const data = await response.json();
          if (Array.isArray(data)) {
            historicalData = data;
            console.log(`Captured data with ${data.length} records`);
          }
        } catch (e) {
          // Ignore parsing errors
        }
      }
    });
    
    // Change timeframe to 1Y to get more historical data
    console.log('Changing timeframe to 1Y...');
    await page.evaluate(() => {
      // Find and click the 1Y button
      const buttons = Array.from(document.querySelectorAll('button'));
      const yearButton = buttons.find(btn => btn.textContent.includes('1Y'));
      if (yearButton) yearButton.click();
    });
    
    // Wait for data to load
    await page.waitForTimeout(5000);
    
    // If we didn't get data from the response listener, try to extract it from the page
    if (historicalData.length === 0) {
      console.log('Attempting to extract data from window object...');
      historicalData = await page.evaluate(() => {
        // Try to find JSON data in the page
        for (const key in window) {
          if (key.includes('__PRELOADED_STATE__')) {
            try {
              const state = window[key];
              if (state && state.index && state.index.price && Array.isArray(state.index.price.priceList)) {
                return state.index.price.priceList;
              }
            } catch (e) {
              // Ignore errors
            }
          }
        }
        return [];
      });
    }
    
    if (historicalData.length > 0) {
      // Test dates
      const now = new Date();
      
      // Find prices for different timeframes
      console.log('\nAnalyzing historical data...');
      
      // Weekly (7 days ago)
      const weeklyDate = new Date(now);
      weeklyDate.setDate(weeklyDate.getDate() - 7);
      
      // Monthly (30 days ago)
      const monthlyDate = new Date(now);
      monthlyDate.setDate(monthlyDate.getDate() - 30);
      
      // YTD (beginning of year)
      const ytdDate = new Date(now.getFullYear(), 0, 2);
      
      // Find closest prices
      const weeklyPrice = findClosestPrice(historicalData, weeklyDate);
      const monthlyPrice = findClosestPrice(historicalData, monthlyDate, 10);
      const ytdPrice = findClosestPrice(historicalData, ytdDate, 15);
      
      // Display results
      console.log('\n--- TEST RESULTS ---');
      console.log(`Weekly (${weeklyDate.toISOString().split('T')[0]}): ${weeklyPrice ? `Success - ${weeklyPrice.date} (${weeklyPrice.price})` : 'Failed'}`);
      console.log(`Monthly (${monthlyDate.toISOString().split('T')[0]}): ${monthlyPrice ? `Success - ${monthlyPrice.date} (${monthlyPrice.price})` : 'Failed'}`);
      console.log(`YTD (${ytdDate.toISOString().split('T')[0]}): ${ytdPrice ? `Success - ${ytdPrice.date} (${ytdPrice.price})` : 'Failed'}`);
      
      if (weeklyPrice && monthlyPrice && ytdPrice) {
        // Calculate percentage changes
        const weeklyChange = ((currentPrice - weeklyPrice.price) / weeklyPrice.price) * 100;
        const monthlyChange = ((currentPrice - monthlyPrice.price) / monthlyPrice.price) * 100;
        const ytdChange = ((currentPrice - ytdPrice.price) / ytdPrice.price) * 100;
        
        console.log('\n--- PERCENTAGE CHANGES ---');
        console.log(`Weekly change: ${weeklyChange.toFixed(2)}%`);
        console.log(`Monthly change: ${monthlyChange.toFixed(2)}%`);
        console.log(`YTD change: ${ytdChange.toFixed(2)}%`);
      }
      
      // Display sample data
      console.log('\n--- SAMPLE DATA ---');
      console.log('First 3 records:');
      historicalData.slice(0, 3).forEach(item => {
        const dateField = item.localTradedAt || item.dt;
        const priceField = item.closePrice || item.ncv;
        console.log(`${dateField}: ${priceField}`);
      });
      
      console.log('\nLast 3 records:');
      historicalData.slice(-3).forEach(item => {
        const dateField = item.localTradedAt || item.dt;
        const priceField = item.closePrice || item.ncv;
        console.log(`${dateField}: ${priceField}`);
      });
    } else {
      console.log('Failed to retrieve historical data');
    }
  } catch (error) {
    console.error('Test failed:', error);
  } finally {
    // Close browser
    await browser.close();
    console.log('\nTest completed.');
  }
}

/**
 * Find closest price to target date
 * @param {Array} data - Historical data
 * @param {Date} targetDate - Target date
 * @param {number} acceptableDiffDays - Acceptable difference in days
 * @returns {Object|null} - Price information or null if not found
 */
function findClosestPrice(data, targetDate, acceptableDiffDays = 7) {
  // Set target date to midnight for comparison
  targetDate.setHours(0, 0, 0, 0);
  
  let closestEntry = null;
  let minDiff = Infinity;
  
  // Iterate through data to find closest date
  for (const entry of data) {
    // Extract date field from either format
    const dateStr = entry.localTradedAt || entry.dt;
    if (!dateStr) continue;
    
    // Parse date
    const entryDate = new Date(dateStr);
    entryDate.setHours(0, 0, 0, 0);
    
    const diff = Math.abs(entryDate.getTime() - targetDate.getTime());
    
    // If exact match, use it
    if (diff === 0) {
      const price = parseFloat(entry.closePrice || entry.ncv);
      console.log(`Found exact match for ${dateStr}: ${price}`);
      return { date: dateStr, price: price, diff: 0 };
    }
    
    // Track closest date
    if (diff < minDiff) {
      minDiff = diff;
      closestEntry = entry;
    }
  }
  
  // Use closest entry if within acceptable range
  if (closestEntry && minDiff <= acceptableDiffDays * 24 * 60 * 60 * 1000) {
    const diffDays = minDiff / (24 * 60 * 60 * 1000);
    const dateStr = closestEntry.localTradedAt || closestEntry.dt;
    const price = parseFloat(closestEntry.closePrice || closestEntry.ncv);
    console.log(`Found closest match for target ${targetDate.toISOString().split('T')[0]}: ${dateStr} (${diffDays.toFixed(1)} days difference): ${price}`);
    return { date: dateStr, price: price, diff: diffDays };
  }
  
  return null;
}

// Run the test
runPlaywrightTest(); 