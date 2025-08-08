import { FileUtils } from '../src/utils';
import * as path from 'path';

// Mock fs-extra to avoid actual filesystem access
jest.mock('fs-extra', () => ({
  pathExists: jest.fn(),
  stat: jest.fn(),
  readFile: jest.fn(),
  ensureDir: jest.fn(),
  writeFile: jest.fn(),
  pathExistsSync: jest.fn(),
  statSync: jest.fn(),
}));

import fs from 'fs-extra';

// Cast to any to simplify typings for mocked methods
const mockedFs = fs as unknown as {
  pathExists: jest.Mock;
  stat: jest.Mock;
  readFile: jest.Mock;
  ensureDir: jest.Mock;
  writeFile: jest.Mock;
  pathExistsSync: jest.Mock;
  statSync: jest.Mock;
};

describe('FileUtils', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('readFile', () => {
    it('reads existing files', async () => {
      mockedFs.pathExists.mockResolvedValue(true);
      mockedFs.stat.mockResolvedValue({ isFile: () => true });
      mockedFs.readFile.mockResolvedValue('content');

      const result = await FileUtils.readFile('/test.txt');
      expect(result).toBe('content');
    });

    it('throws for missing files', async () => {
      mockedFs.pathExists.mockResolvedValue(false);
      await expect(FileUtils.readFile('/missing.txt')).rejects.toThrow('File not found');
    });
  });

  describe('writeFile', () => {
    it('ensures directory and writes file', async () => {
      mockedFs.ensureDir.mockResolvedValue(undefined);
      mockedFs.writeFile.mockResolvedValue(undefined);

      const target = path.join('/dir', 'file.txt');
      await FileUtils.writeFile(target, 'data');

      expect(mockedFs.ensureDir).toHaveBeenCalledWith('/dir');
      expect(mockedFs.writeFile).toHaveBeenCalledWith(target, 'data');
    });
  });

  describe('validateFile', () => {
    it('detects missing files', () => {
      mockedFs.pathExistsSync.mockReturnValue(false);
      const result = FileUtils.validateFile('/missing.txt', ['.txt']);
      expect(result.isValid).toBe(false);
      expect(result.errors[0].code).toBe('FILE_NOT_FOUND');
    });

    it('warns about large files', () => {
      mockedFs.pathExistsSync.mockReturnValue(true);
      mockedFs.statSync.mockReturnValue({ size: 60 * 1024 * 1024 });
      const result = FileUtils.validateFile('/file.txt', ['.txt']);
      expect(result.isValid).toBe(true);
      expect(result.warnings[0].code).toBe('LARGE_FILE_SIZE');
    });
  });

  it('generates unique filenames', () => {
    const name = FileUtils.generateUniqueFileName('report.docx');
    expect(name).toMatch(/^report_[^_]{8}\.docx$/);
  });

  it('sanitizes filenames', () => {
    const name = FileUtils.sanitizeFileName('in*valid:file?.txt');
    expect(name).toBe('in_valid_file_.txt');
  });
});
