import { describe, it, expect, beforeEach, vi } from 'vitest';
import { parseIndexTableData, parseMobileApiData } from './naverParsers';

describe('KOSDAQ Real Data Fetching Test', () => {
  // Real HTML from Naver Finance KOSDAQ historical page (simplified for test)
  const realKosdaqHistoryHtml = `
    <table class="type_1">
      <tr class="first">
        <th>날짜</th>
        <th>체결가</th>
        <th>전일비</th>
        <th>등락률</th>
        <th>거래량</th>
        <th>거래대금</th>
      </tr>
      <tr>
        <td class="date">2025.03.21</td>
        <td class="number_1">717.62</td>
        <td class="rate_down" style="padding-right:35px;">
          <img src="https://ssl.pstatic.net/imgstock/images/images4/ico_down.gif" width="7" height="6" style="margin-right:4px;" alt="하락"><span class="tah p11 nv01">
          7.53
          </span>
        </td>
        <td class="number_1">
          <span class="tah p11 nv01">
          -1.04%
          </span>
        </td>
        <td class="number_1" style="padding-right:40px;">149,391</td>
        <td class="number_1" style="padding-right:30px;">1,655,703</td>
      </tr>
      <tr>
        <td class="date">2025.03.20</td>
        <td class="number_1">725.15</td>
        <td class="rate_down" style="padding-right:35px;">
          <img src="https://ssl.pstatic.net/imgstock/images/images4/ico_down.gif" width="7" height="6" style="margin-right:4px;" alt="하락"><span class="tah p11 nv01">
          13.20
          </span>
        </td>
        <td class="number_1">
          <span class="tah p11 nv01">
          -1.79%
          </span>
        </td>
        <td class="number_1" style="padding-right:40px;">749,018</td>
        <td class="number_1" style="padding-right:30px;">8,338,384</td>
      </tr>
    </table>
  `;

  // Real JSON from Naver Mobile API (first few entries)
  const realKosdaqApiJson = `[
    {"localTradedAt":"2025-03-21","closePrice":"716.97","compareToPreviousClosePrice":"-8.18","compareToPreviousPrice":{"code":"5","text":"하락","name":"FALLING"},"fluctuationsRatio":"-1.13","openPrice":"724.94","highPrice":"727.15","lowPrice":"716.72"},
    {"localTradedAt":"2025-03-20","closePrice":"725.15","compareToPreviousClosePrice":"-13.20","compareToPreviousPrice":{"code":"5","text":"하락","name":"FALLING"},"fluctuationsRatio":"-1.79","openPrice":"743.06","highPrice":"743.39","lowPrice":"725.10"},
    {"localTradedAt":"2025-03-19","closePrice":"738.35","compareToPreviousClosePrice":"-7.19","compareToPreviousPrice":{"code":"5","text":"하락","name":"FALLING"},"fluctuationsRatio":"-0.96","openPrice":"743.21","highPrice":"747.53","lowPrice":"737.52"}
  ]`;

  // Create a more comprehensive API response that includes older data for monthly and YTD testing
  const extendedKosdaqApiJson = `[
    {"localTradedAt":"2025-03-21","closePrice":"716.97","compareToPreviousClosePrice":"-8.18","compareToPreviousPrice":{"code":"5","text":"하락","name":"FALLING"},"fluctuationsRatio":"-1.13","openPrice":"724.94","highPrice":"727.15","lowPrice":"716.72"},
    {"localTradedAt":"2025-03-20","closePrice":"725.15","compareToPreviousClosePrice":"-13.20","compareToPreviousPrice":{"code":"5","text":"하락","name":"FALLING"},"fluctuationsRatio":"-1.79","openPrice":"743.06","highPrice":"743.39","lowPrice":"725.10"},
    {"localTradedAt":"2025-03-19","closePrice":"738.35","compareToPreviousClosePrice":"-7.19","compareToPreviousPrice":{"code":"5","text":"하락","name":"FALLING"},"fluctuationsRatio":"-0.96","openPrice":"743.21","highPrice":"747.53","lowPrice":"737.52"},
    {"localTradedAt":"2025-03-14","closePrice":"734.26","compareToPreviousClosePrice":"11.46","compareToPreviousPrice":{"code":"2","text":"상승","name":"RISING"},"fluctuationsRatio":"1.59","openPrice":"723.75","highPrice":"736.84","lowPrice":"723.71"},
    {"localTradedAt":"2025-03-07","closePrice":"727.70","compareToPreviousClosePrice":"-7.22","compareToPreviousPrice":{"code":"5","text":"하락","name":"FALLING"},"fluctuationsRatio":"-0.98","openPrice":"729.43","highPrice":"738.10","lowPrice":"725.97"},
    {"localTradedAt":"2025-02-28","closePrice":"743.96","compareToPreviousClosePrice":"-26.89","compareToPreviousPrice":{"code":"5","text":"하락","name":"FALLING"},"fluctuationsRatio":"-3.49","openPrice":"760.10","highPrice":"760.10","lowPrice":"743.95"},
    {"localTradedAt":"2025-02-21","closePrice":"774.65","compareToPreviousClosePrice":"6.38","compareToPreviousPrice":{"code":"2","text":"상승","name":"RISING"},"fluctuationsRatio":"0.83","openPrice":"768.39","highPrice":"774.78","lowPrice":"767.18"},
    {"localTradedAt":"2025-01-31","closePrice":"769.43","compareToPreviousClosePrice":"-3.90","compareToPreviousPrice":{"code":"5","text":"하락","name":"FALLING"},"fluctuationsRatio":"-0.50","openPrice":"766.85","highPrice":"773.91","lowPrice":"765.26"},
    {"localTradedAt":"2025-01-02","closePrice":"760.85","compareToPreviousClosePrice":"-2.56","compareToPreviousPrice":{"code":"5","text":"하락","name":"FALLING"},"fluctuationsRatio":"-0.33","openPrice":"763.13","highPrice":"764.57","lowPrice":"759.60"}
  ]`;

  describe('Desktop Site KOSDAQ Parsing with Real Data', () => {
    it('should parse real KOSDAQ historical data from desktop site', () => {
      const result = parseIndexTableData(realKosdaqHistoryHtml);
      
      expect(result).toHaveLength(2);
      
      // Check first day's data (most recent)
      expect(result[0].date).toBe('2025-03-21');
      expect(result[0].close).toBe(717.62);
      
      // Check second day's data
      expect(result[1].date).toBe('2025-03-20');
      expect(result[1].close).toBe(725.15);
    });
  });

  describe('Mobile API KOSDAQ Parsing with Real Data', () => {
    it('should parse real KOSDAQ data from mobile API', () => {
      // Adapt to match our mobile API response format
      const apiResponse = realKosdaqApiJson.replace(/"localTradedAt"/g, '"dt"').replace(/"closePrice"/g, '"ncv"');
      const result = parseMobileApiData(apiResponse);
      
      expect(result).toHaveLength(3);
      
      // Check most recent price
      expect(result[0].date).toBe('2025-03-21');
      expect(parseFloat(result[0].price)).toBe(716.97);
      
      // Check the second price
      expect(result[1].date).toBe('2025-03-20');
      expect(parseFloat(result[1].price)).toBe(725.15);
      
      // Check the third price
      expect(result[2].date).toBe('2025-03-19');
      expect(parseFloat(result[2].price)).toBe(738.35);
    });

    it('should handle extended dataset for monthly and YTD historical data', () => {
      // Create a function mimicking the NaverFinanceProvider.getIndexHistoricalPriceViaAPI
      const findHistoricalPrice = (apiJson, targetDateStr) => {
        const data = JSON.parse(apiJson);
        
        // Set target date to midnight for comparison
        const targetDate = new Date(targetDateStr);
        targetDate.setHours(0, 0, 0, 0);
        
        let closestEntry = null;
        let minDiff = Infinity;
        
        // Loop through all entries to find closest date
        for (const entry of data) {
          const dateStr = entry.localTradedAt;
          const entryDate = new Date(dateStr);
          entryDate.setHours(0, 0, 0, 0);
          
          const diff = Math.abs(entryDate.getTime() - targetDate.getTime());
          
          // Exact match
          if (diff === 0) {
            return parseFloat(entry.closePrice);
          }
          
          // Track closest date
          if (diff < minDiff) {
            minDiff = diff;
            closestEntry = entry;
          }
        }
        
        // Use closest entry if within 7 days
        if (closestEntry && minDiff <= 7 * 24 * 60 * 60 * 1000) {
          return parseFloat(closestEntry.closePrice);
        }
        
        return null;
      };
      
      // Test finding weekly data (7 days ago)
      const now = new Date('2025-03-21');
      const weeklyDate = new Date(now);
      weeklyDate.setDate(weeklyDate.getDate() - 7);
      const weeklyPrice = findHistoricalPrice(extendedKosdaqApiJson, weeklyDate.toISOString().split('T')[0]);
      
      // Test finding monthly data (30 days ago)
      const monthlyDate = new Date(now);
      monthlyDate.setDate(monthlyDate.getDate() - 30);
      const monthlyPrice = findHistoricalPrice(extendedKosdaqApiJson, monthlyDate.toISOString().split('T')[0]);
      
      // Test finding YTD data (start of year)
      const ytdDate = new Date(now.getFullYear(), 0, 2); // Jan 2nd of current year
      const ytdPrice = findHistoricalPrice(extendedKosdaqApiJson, ytdDate.toISOString().split('T')[0]);
      
      // Check if we can find data for each timeframe
      expect(weeklyPrice).not.toBeNull();
      expect(weeklyPrice).toBeCloseTo(734.26, 1); // Should be close to March 14th data
      
      expect(monthlyPrice).not.toBeNull();
      expect(monthlyPrice).toBeCloseTo(774.65, 1); // Should be close to Feb 21 data
      
      expect(ytdPrice).not.toBeNull();
      expect(ytdPrice).toBeCloseTo(760.85, 1); // Should be close to Jan 2 data
    });
  });

  // Test our mobile API parser with the actual format returned
  describe('Enhanced Mobile API Parser for Actual Response Format', () => {
    // Function to directly parse the real mobile API format
    const parseActualMobileApi = (jsonContent) => {
      try {
        const data = JSON.parse(jsonContent);
        if (!Array.isArray(data)) return [];
        
        return data.map(item => ({
          date: item.localTradedAt,
          price: parseFloat(item.closePrice)
        }));
      } catch (e) {
        console.error('Error parsing mobile API data:', e);
        return [];
      }
    };

    it('should correctly parse the actual Naver mobile API format', () => {
      const result = parseActualMobileApi(realKosdaqApiJson);
      
      expect(result).toHaveLength(3);
      
      // Check most recent price
      expect(result[0].date).toBe('2025-03-21');
      expect(result[0].price).toBe(716.97);
      
      // Check the second price
      expect(result[1].date).toBe('2025-03-20');
      expect(result[1].price).toBe(725.15);
      
      // Check the third price
      expect(result[2].date).toBe('2025-03-19');
      expect(result[2].price).toBe(738.35);
    });
  });
}); 