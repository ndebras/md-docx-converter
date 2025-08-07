// Jest setup file
import 'jest';

// Extend Jest matchers
expect.extend({
  toBeValidMarkdown(received: string) {
    const pass = received.trim().length > 0 && !received.includes('<html>');
    return {
      message: () => `expected ${received} to be valid Markdown`,
      pass,
    };
  },
  
  toBeValidDocx(received: Buffer) {
    const pass = received.length > 0 && received.subarray(0, 4).toString('hex') === '504b0304';
    return {
      message: () => `expected buffer to be valid DOCX format`,
      pass,
    };
  },
});

// Global test timeout
jest.setTimeout(30000);

// Mock console methods in tests
beforeEach(() => {
  jest.spyOn(console, 'log').mockImplementation(() => {});
  jest.spyOn(console, 'warn').mockImplementation(() => {});
  jest.spyOn(console, 'error').mockImplementation(() => {});
});

afterEach(() => {
  jest.restoreAllMocks();
});
