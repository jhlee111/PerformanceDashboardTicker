/**
 * Extract and parse price from Naver Finance HTML content
 * @param {string} content - HTML content
 * @return {number|null} Price or null if not found
 */
export function parsePrice(content) {
  // Find the current price (today's close or latest price)
  // The price is in a tag with ID "_nowVal"
  const priceRegex = /<span id="_nowVal"[^>]*>([\d,]+)<\/span>/;
  const priceMatch = content.match(priceRegex);
  
  if (priceMatch && priceMatch[1]) {
    // Remove commas and convert to number
    return parseFloat(priceMatch[1].replace(/,/g, ''));
  }
  
  // If we can't find the price with the first pattern, try another one
  const altPriceRegex = /<dd class="no_today">[\s\n]*<span class="no_up">[\s\n]*<span class="blind">현재가<\/span>[\s\n]*<span class="no_up">[\s\n]*([\d,]+)/;
  const altPriceMatch = content.match(altPriceRegex);
  
  if (altPriceMatch && altPriceMatch[1]) {
    // Remove commas and convert to number
    return parseFloat(altPriceMatch[1].replace(/,/g, ''));
  }
  
  // If we still can't find the price, look for the table with recent prices
  const tableRegex = /<td class="num">([\d,]+)<\/td>/;
  const tableMatch = content.match(tableRegex);
  
  if (tableMatch && tableMatch[1]) {
    // Remove commas and convert to number
    return parseFloat(tableMatch[1].replace(/,/g, ''));
  }
  
  return null;
}

/**
 * Extract and parse index price from Naver Finance HTML content
 * @param {string} content - HTML content
 * @return {number|null} Index price or null if not found
 */
export function parseIndexPrice(content) {
  // Try multiple patterns for finding the index value
  // Pattern 1: Standard pattern in newer pages
  const indexRegex1 = /<span class="num">([\d,.]+)<\/span>/;
  const indexMatch1 = content.match(indexRegex1);
  
  if (indexMatch1 && indexMatch1[1]) {
    // Remove commas and convert to number
    return parseFloat(indexMatch1[1].replace(/,/g, ''));
  }
  
  // Pattern 2: Another possible location
  const indexRegex2 = /<em id="_nowVal"[^>]*>([\d,.]+)<\/em>/;
  const indexMatch2 = content.match(indexRegex2);
  
  if (indexMatch2 && indexMatch2[1]) {
    // Remove commas and convert to number
    return parseFloat(indexMatch2[1].replace(/,/g, ''));
  }
  
  // Pattern 3: Try a more general approach to find any number that might be the index
  const indexRegex3 = /<em[^>]*>([\d,.]+)<\/em>/;
  const indexMatch3 = content.match(indexRegex3);
  
  if (indexMatch3 && indexMatch3[1]) {
    // Remove commas and convert to number
    return parseFloat(indexMatch3[1].replace(/,/g, ''));
  }
  
  // If all else fails, try getting it from the title
  const titleRegex = /<title>.*?([0-9,.]+).*?<\/title>/;
  const titleMatch = content.match(titleRegex);
  
  if (titleMatch && titleMatch[1]) {
    // Remove commas and convert to number
    return parseFloat(titleMatch[1].replace(/,/g, ''));
  }
  
  return null;
}

/**
 * Parse HTML table data into structured format for regular stocks
 * @param {string} content - HTML content
 * @return {Array} Array of objects with date and close price
 */
export function parseTableData(content) {
  const result = [];
  
  // Regular expression to find table rows with date and close price
  const rowRegex = /<tr[^>]*>[\s\S]*?<td[^>]*>[\s\S]*?<span[^>]*>(\d{4}.\d{2}.\d{2})[\s\S]*?<\/span>[\s\S]*?<\/td>[\s\S]*?<td[^>]*>[\s\S]*?<span[^>]*>([\d,]+)<\/span>[\s\S]*?<\/td>/g;
  
  let match;
  while ((match = rowRegex.exec(content)) !== null) {
    const dateStr = match[1].replace(/\./g, '-');
    const closePrice = parseFloat(match[2].replace(/,/g, ''));
    
    result.push({
      date: dateStr,
      close: closePrice
    });
  }
  
  return result;
}

/**
 * Parse index table data into structured format for KOSPI/KOSDAQ
 * @param {string} content - HTML content
 * @return {Array} Array of objects with date and close price
 */
export function parseIndexTableData(content) {
  const result = [];
  
  // Regular expression to find date and close price in index pages
  // Updated pattern to match real Naver Finance HTML structure
  const rowRegex = /<tr[^>]*>[\s\S]*?<td[^>]*class="date">(\d{4}.\d{2}.\d{2})[\s\S]*?<\/td>[\s\S]*?<td[^>]*class="number_1">([\d,.]+)<\/td>/g;
  
  let match;
  while ((match = rowRegex.exec(content)) !== null) {
    const dateStr = match[1].replace(/\./g, '-');
    const closePrice = parseFloat(match[2].replace(/,/g, ''));
    
    result.push({
      date: dateStr,
      close: closePrice
    });
  }
  
  return result;
}

/**
 * Parse mobile web stock page content to extract JSON data
 * @param {string} content - HTML content from mobile stock page
 * @return {object|null} Parsed JSON object or null if not found
 */
export function parseMobileStockData(content) {
  // Look for embedded JSON data in the page
  const jsonRegex = /window\.__PRELOADED_STATE__\s*=\s*({.*?});/s;
  const match = content.match(jsonRegex);
  
  if (match && match[1]) {
    try {
      return JSON.parse(match[1]);
    } catch (e) {
      console.error('Failed to parse JSON from mobile page:', e);
    }
  }
  
  // Alternative pattern that might appear in some pages
  const dataRegex = /"symbolCode":"([^"]*?)","name":"([^"]*?)","price":([\d.]+)/;
  const dataMatch = content.match(dataRegex);
  
  if (dataMatch) {
    return {
      symbolCode: dataMatch[1],
      name: dataMatch[2],
      price: parseFloat(dataMatch[3])
    };
  }
  
  return null;
}

/**
 * Parse mobile API response for index historical data
 * @param {string} content - JSON content from mobile API
 * @return {Array} Array of parsed price data objects
 */
export function parseMobileApiData(content) {
  try {
    const jsonData = JSON.parse(content);
    
    if (!Array.isArray(jsonData)) {
      return [];
    }
    
    return jsonData.map(entry => {
      // Handle both expected API format and actual API format
      if (entry.dt && entry.ncv) {
        // Original format with dt and ncv fields
        return {
          date: entry.dt,
          price: typeof entry.ncv === 'string' ? parseFloat(entry.ncv) : entry.ncv
        };
      } else if (entry.localTradedAt && entry.closePrice) {
        // Actual Naver mobile API format
        return {
          date: entry.localTradedAt,
          price: typeof entry.closePrice === 'string' ? parseFloat(entry.closePrice) : entry.closePrice
        };
      } else {
        // Fallback for other formats
        const dateField = entry.dt || entry.localTradedAt || '';
        const priceField = entry.ncv || entry.closePrice || null;
        return {
          date: dateField,
          price: typeof priceField === 'string' ? parseFloat(priceField) : priceField
        };
      }
    });
  } catch (e) {
    console.error('Failed to parse mobile API data:', e);
    return [];
  }
} 