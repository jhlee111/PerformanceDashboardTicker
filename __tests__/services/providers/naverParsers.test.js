import { describe, it, expect } from 'vitest';
import { 
  parsePrice,
  parseIndexPrice,
  parseTableData,
  parseIndexTableData,
  parseMobileStockData,
  parseMobileApiData
} from './naverParsers';

describe('Naver Finance Parsers', () => {
  describe('parsePrice', () => {
    it('should extract price with _nowVal pattern', () => {
      const html = `
        <div class="stock_info">
          <span id="_nowVal" class="price">32,500</span>
        </div>
      `;
      expect(parsePrice(html)).toBe(32500);
    });

    it('should extract price with alternative pattern', () => {
      const html = `
        <dd class="no_today">
          <span class="no_up">
            <span class="blind">현재가</span>
            <span class="no_up">45,670</span>
          </span>
        </dd>
      `;
      expect(parsePrice(html)).toBe(45670);
    });

    it('should extract price from table as last resort', () => {
      const html = `
        <table>
          <tr>
            <td class="num">12,345</td>
          </tr>
        </table>
      `;
      expect(parsePrice(html)).toBe(12345);
    });

    it('should return null when no price is found', () => {
      const html = '<div>No price here</div>';
      expect(parsePrice(html)).toBeNull();
    });
  });

  describe('parseIndexPrice', () => {
    it('should extract index price with span.num pattern', () => {
      const html = '<div><span class="num">2,718.28</span></div>';
      expect(parseIndexPrice(html)).toBe(2718.28);
    });

    it('should extract index price with _nowVal pattern', () => {
      const html = '<div><em id="_nowVal">3,141.59</em></div>';
      expect(parseIndexPrice(html)).toBe(3141.59);
    });

    it('should extract index price with any em tag as fallback', () => {
      const html = '<div><em>718.87</em></div>';
      expect(parseIndexPrice(html)).toBe(718.87);
    });

    it('should extract index price from title as last resort', () => {
      const html = '<title>KOSDAQ 782.33 | 네이버 금융</title>';
      expect(parseIndexPrice(html)).toBe(782.33);
    });

    it('should return null when no index price is found', () => {
      const html = '<div>No index data</div>';
      expect(parseIndexPrice(html)).toBeNull();
    });
  });

  describe('parseTableData', () => {
    it('should parse table data for regular stocks', () => {
      const html = `
        <table>
          <tr>
            <td><span>2025.03.21</span></td>
            <td><span>42,500</span></td>
          </tr>
          <tr>
            <td><span>2025.03.20</span></td>
            <td><span>42,100</span></td>
          </tr>
        </table>
      `;
      const result = parseTableData(html);
      expect(result).toHaveLength(2);
      expect(result[0].date).toBe('2025-03-21');
      expect(result[0].close).toBe(42500);
      expect(result[1].date).toBe('2025-03-20');
      expect(result[1].close).toBe(42100);
    });

    it('should return empty array when no table data is found', () => {
      const html = '<div>No table here</div>';
      expect(parseTableData(html)).toEqual([]);
    });
  });

  describe('parseIndexTableData', () => {
    it('should parse index table data', () => {
      const html = `
        <table>
          <tr>
            <td>2025.03.21</td>
            <td><span>2,718.28</span></td>
          </tr>
          <tr>
            <td>2025.03.20</td>
            <td><span>2,715.33</span></td>
          </tr>
        </table>
      `;
      const result = parseIndexTableData(html);
      expect(result).toHaveLength(2);
      expect(result[0].date).toBe('2025-03-21');
      expect(result[0].close).toBe(2718.28);
      expect(result[1].date).toBe('2025-03-20');
      expect(result[1].close).toBe(2715.33);
    });

    it('should return empty array when no index table data is found', () => {
      const html = '<div>No index table here</div>';
      expect(parseIndexTableData(html)).toEqual([]);
    });
  });

  describe('parseMobileStockData', () => {
    it('should parse embedded JSON data', () => {
      const html = `
        <script>
          window.__PRELOADED_STATE__ = {"symbolCode":"KOSDAQ","name":"코스닥","price":718.87};
        </script>
      `;
      const result = parseMobileStockData(html);
      expect(result).toEqual({
        symbolCode: 'KOSDAQ',
        name: '코스닥',
        price: 718.87
      });
    });

    it('should parse data using regex as fallback', () => {
      const html = `
        <div data-item="symbolCode":"KOSDAQ","name":"코스닥","price":718.87</div>
      `;
      const result = parseMobileStockData(html);
      expect(result).toEqual({
        symbolCode: 'KOSDAQ',
        name: '코스닥',
        price: 718.87
      });
    });

    it('should return null when no data is found', () => {
      const html = '<div>No mobile data</div>';
      expect(parseMobileStockData(html)).toBeNull();
    });
  });

  describe('parseMobileApiData', () => {
    it('should parse mobile API JSON data', () => {
      const json = `[
        {"dt":"2025.03.21","ncv":718.87},
        {"dt":"2025.03.20","ncv":717.33}
      ]`;
      const result = parseMobileApiData(json);
      expect(result).toHaveLength(2);
      expect(result[0]).toEqual({ date: '2025.03.21', price: 718.87 });
      expect(result[1]).toEqual({ date: '2025.03.20', price: 717.33 });
    });

    it('should handle invalid JSON', () => {
      const json = 'Not valid JSON';
      expect(parseMobileApiData(json)).toEqual([]);
    });

    it('should handle non-array JSON', () => {
      const json = '{"error": "Not found"}';
      expect(parseMobileApiData(json)).toEqual([]);
    });
  });
}); 