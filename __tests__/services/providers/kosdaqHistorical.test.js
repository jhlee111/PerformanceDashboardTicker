import { describe, it, expect, beforeEach, vi } from 'vitest';
import { parseIndexTableData, parseMobileApiData } from './naverParsers';

// This test file focuses specifically on KOSDAQ historical data parsing
// since that's where we're having issues

describe('KOSDAQ Historical Data Parsing', () => {
  // Sample HTML from desktop site with KOSDAQ historical data
  const sampleDesktopHtml = `
    <table class="type_1">
      <tr class="first">
        <th>날짜</th>
        <th>종가</th>
        <th>전일비</th>
        <th>시가</th>
        <th>고가</th>
        <th>저가</th>
        <th>거래량</th>
      </tr>
      <tr>
        <td>2025.03.21</td>
        <td><span class="tah p11">718.87</span></td>
        <td><span class="tah p11 red">+2.09</span></td>
        <td><span class="tah p11">717.33</span></td>
        <td><span class="tah p11">719.54</span></td>
        <td><span class="tah p11">716.21</span></td>
        <td><span class="tah p11">532,435</span></td>
      </tr>
      <tr>
        <td>2025.03.20</td>
        <td><span class="tah p11">716.78</span></td>
        <td><span class="tah p11 blue">-1.22</span></td>
        <td><span class="tah p11">717.45</span></td>
        <td><span class="tah p11">718.33</span></td>
        <td><span class="tah p11">715.67</span></td>
        <td><span class="tah p11">487,651</span></td>
      </tr>
      <tr>
        <td>2025.03.19</td>
        <td><span class="tah p11">718.00</span></td>
        <td><span class="tah p11 red">+0.85</span></td>
        <td><span class="tah p11">717.54</span></td>
        <td><span class="tah p11">718.92</span></td>
        <td><span class="tah p11">716.43</span></td>
        <td><span class="tah p11">498,742</span></td>
      </tr>
    </table>
  `;

  // Sample mobile API JSON data for KOSDAQ
  const sampleMobileApiJson = `[
    {
      "dt": "2025.03.21",
      "ncv": 718.87,
      "cr": 0.29,
      "cv": 2.09,
      "ov": 717.33,
      "hv": 719.54,
      "lv": 716.21,
      "tv": 532435
    },
    {
      "dt": "2025.03.20",
      "ncv": 716.78,
      "cr": -0.17,
      "cv": -1.22,
      "ov": 717.45,
      "hv": 718.33,
      "lv": 715.67,
      "tv": 487651
    },
    {
      "dt": "2025.03.19",
      "ncv": 718.00,
      "cr": 0.12,
      "cv": 0.85,
      "ov": 717.54,
      "hv": 718.92,
      "lv": 716.43,
      "tv": 498742
    }
  ]`;

  describe('Desktop Site KOSDAQ Historical Parsing', () => {
    it('should parse KOSDAQ historical data from desktop site', () => {
      const result = parseIndexTableData(sampleDesktopHtml);
      
      expect(result).toHaveLength(3);
      
      // Check first day's data
      expect(result[0].date).toBe('2025-03-21');
      expect(result[0].close).toBe(718.87);
      
      // Check second day's data
      expect(result[1].date).toBe('2025-03-20');
      expect(result[1].close).toBe(716.78);
      
      // Check third day's data
      expect(result[2].date).toBe('2025-03-19');
      expect(result[2].close).toBe(718.00);
    });

    it('should handle empty or invalid HTML', () => {
      const result = parseIndexTableData('<div>No table data</div>');
      expect(result).toEqual([]);
    });
  });

  describe('Mobile API KOSDAQ Historical Parsing', () => {
    it('should parse KOSDAQ historical data from mobile API', () => {
      const result = parseMobileApiData(sampleMobileApiJson);
      
      expect(result).toHaveLength(3);
      
      // Check first day's data
      expect(result[0].date).toBe('2025.03.21');
      expect(result[0].price).toBe(718.87);
      
      // Check second day's data
      expect(result[1].date).toBe('2025.03.20');
      expect(result[1].price).toBe(716.78);
      
      // Check third day's data
      expect(result[2].date).toBe('2025.03.19');
      expect(result[2].price).toBe(718.00);
    });

    it('should handle empty array', () => {
      const result = parseMobileApiData('[]');
      expect(result).toEqual([]);
    });

    it('should handle malformed JSON', () => {
      const result = parseMobileApiData('{malformed:json}');
      expect(result).toEqual([]);
    });
  });
}); 