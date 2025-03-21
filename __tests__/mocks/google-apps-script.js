// Mock Google Apps Script's global objects and functions

// Mock Logger
global.Logger = {
  log: (message) => console.log(message),
};

// Mock Utilities
global.Utilities = {
  formatDate: (date, timezone, format) => {
    const d = new Date(date);
    if (format === 'yyyyMMdd') {
      return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
    }
    if (format === 'yyyy-MM-dd') {
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    }
    return d.toISOString();
  },
};

// Mock UrlFetchApp
global.UrlFetchApp = {
  fetch: vi.fn(),
};

// Mock response object
class HttpResponse {
  constructor(code, content) {
    this.code = code;
    this.content = content;
  }
  
  getResponseCode() {
    return this.code;
  }
  
  getContentText() {
    return this.content;
  }
}

// Function to create mock HTTP responses
global.mockHttpResponse = (code, content) => {
  return new HttpResponse(code, content);
};

// Mock CONFIG object
global.CONFIG = {
  STATUS: {
    NO_DATA: null,
  },
};

// Mock SS global
global.SS = {
  getSheetByName: vi.fn().mockReturnValue({
    getRange: vi.fn().mockReturnValue({
      setValue: vi.fn(),
    }),
  }),
  insertSheet: vi.fn().mockReturnValue({
    getRange: vi.fn().mockReturnValue({
      setValue: vi.fn(),
    }),
  }),
};

// Mock diagnostic function
global.isDiagnosticModeEnabled = vi.fn().mockReturnValue(false);

export { HttpResponse }; 