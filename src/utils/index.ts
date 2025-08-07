import * as fs from 'fs-extra';
import * as path from 'path';
import { v4 as uuidv4 } from 'uuid';
import * as mime from 'mime-types';
import { ValidationResult, ValidationError, ValidationWarning } from '../types';
import { logger } from './logger';

/**
 * File system utilities with enhanced error handling
 */
export class FileUtils {
  /**
   * Safely read file with encoding detection
   */
  static async readFile(filePath: string): Promise<string> {
    try {
      const exists = await fs.pathExists(filePath);
      if (!exists) {
        throw new Error(`File not found: ${filePath}`);
      }

      const stats = await fs.stat(filePath);
      if (!stats.isFile()) {
        throw new Error(`Path is not a file: ${filePath}`);
      }

      return await fs.readFile(filePath, 'utf-8');
    } catch (error) {
      logger.error(`Failed to read file: ${filePath}`, { error });
      throw error;
    }
  }

  /**
   * Safely write file with directory creation
   */
  static async writeFile(filePath: string, content: string | Buffer): Promise<void> {
    try {
      const dir = path.dirname(filePath);
      await fs.ensureDir(dir);
      await fs.writeFile(filePath, content);
      logger.debug(`File written successfully: ${filePath}`);
    } catch (error) {
      logger.error(`Failed to write file: ${filePath}`, { error });
      throw error;
    }
  }

  /**
   * Get file extension
   */
  static getExtension(filePath: string): string {
    return path.extname(filePath).toLowerCase();
  }

  /**
   * Get MIME type from file extension
   */
  static getMimeType(filePath: string): string | false {
    return mime.lookup(filePath);
  }

  /**
   * Validate file path and type
   */
  static validateFile(filePath: string, allowedExtensions: string[]): ValidationResult {
    const errors: ValidationError[] = [];
    const warnings: ValidationWarning[] = [];

    // Check if file exists
    if (!fs.pathExistsSync(filePath)) {
      errors.push({
        code: 'FILE_NOT_FOUND',
        message: `File not found: ${filePath}`,
      });
      return { isValid: false, errors, warnings };
    }

    // Check file extension
    const ext = this.getExtension(filePath);
    if (!allowedExtensions.includes(ext)) {
      errors.push({
        code: 'INVALID_FILE_TYPE',
        message: `Invalid file type: ${ext}. Allowed types: ${allowedExtensions.join(', ')}`,
      });
    }

    // Check file size (warn if > 50MB)
    const stats = fs.statSync(filePath);
    const fileSizeMB = stats.size / (1024 * 1024);
    if (fileSizeMB > 50) {
      warnings.push({
        code: 'LARGE_FILE_SIZE',
        message: `Large file detected: ${fileSizeMB.toFixed(2)}MB. Processing may be slow.`,
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
    };
  }

  /**
   * Generate unique filename
   */
  static generateUniqueFileName(originalName: string, suffix?: string): string {
    const ext = path.extname(originalName);
    const baseName = path.basename(originalName, ext);
    const uniqueId = uuidv4().slice(0, 8);
    const suffixPart = suffix ? `_${suffix}` : '';
    return `${baseName}${suffixPart}_${uniqueId}${ext}`;
  }

  /**
   * Sanitize filename for cross-platform compatibility
   */
  static sanitizeFileName(fileName: string): string {
    // Remove or replace invalid characters
    return fileName
      .replace(/[<>:"/\\|?*]/g, '_')
      .replace(/\s+/g, '_')
      .replace(/_+/g, '_')
      .replace(/^_|_$/g, '');
  }

  /**
   * Get relative path for output
   */
  static getRelativePath(from: string, to: string): string {
    return path.relative(from, to).replace(/\\/g, '/');
  }

  /**
   * Ensure directory exists and is writable
   */
  static async ensureWritableDir(dirPath: string): Promise<void> {
    try {
      await fs.ensureDir(dirPath);
      
      // Test write permissions
      const testFile = path.join(dirPath, '.write_test');
      await fs.writeFile(testFile, 'test');
      await fs.remove(testFile);
      
      logger.debug(`Directory ready: ${dirPath}`);
    } catch (error) {
      logger.error(`Failed to ensure writable directory: ${dirPath}`, { error });
      throw new Error(`Cannot write to directory: ${dirPath}`);
    }
  }
}

/**
 * String utilities for text processing
 */
export class StringUtils {
  /**
   * Generate URL-safe slug from text
   */
  static slugify(text: string): string {
    return text
      .toLowerCase()
      .trim()
      .replace(/[^\w\s-]/g, '')
      .replace(/[\s_-]+/g, '-')
      .replace(/^-+|-+$/g, '');
  }

  /**
   * Extract heading anchor from text
   */
  static createAnchor(heading: string): string {
    return this.slugify(heading);
  }

  /**
   * Escape HTML special characters
   */
  static escapeHtml(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  /**
   * Unescape HTML entities
   */
  static unescapeHtml(text: string): string {
    return text
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'");
  }

  /**
   * Clean whitespace and normalize line endings
   */
  static normalizeWhitespace(text: string): string {
    return text
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .replace(/[ \t]+$/gm, '')
      .replace(/\n{3,}/g, '\n\n');
  }

  /**
   * Word wrap text to specified width
   */
  static wordWrap(text: string, width: number = 80): string {
    const lines = text.split('\n');
    const wrappedLines: string[] = [];

    for (const line of lines) {
      if (line.length <= width) {
        wrappedLines.push(line);
        continue;
      }

      const words = line.split(' ');
      let currentLine = '';

      for (const word of words) {
        if ((currentLine + word).length <= width) {
          currentLine += (currentLine ? ' ' : '') + word;
        } else {
          if (currentLine) {
            wrappedLines.push(currentLine);
          }
          currentLine = word;
        }
      }

      if (currentLine) {
        wrappedLines.push(currentLine);
      }
    }

    return wrappedLines.join('\n');
  }

  /**
   * Extract text content from HTML
   */
  static stripHtml(html: string): string {
    return html.replace(/<[^>]*>/g, '').trim();
  }

  /**
   * Count words in text
   */
  static countWords(text: string): number {
    return text.trim().split(/\s+/).filter(word => word.length > 0).length;
  }

  /**
   * Estimate reading time in minutes
   */
  static estimateReadingTime(text: string, wordsPerMinute: number = 200): number {
    const wordCount = this.countWords(text);
    return Math.ceil(wordCount / wordsPerMinute);
  }
}

/**
 * Performance utilities
 */
export class PerformanceUtils {
  private static timers: Map<string, number> = new Map();

  /**
   * Start performance timer
   */
  static startTimer(label: string): void {
    this.timers.set(label, Date.now());
  }

  /**
   * End performance timer and return elapsed time
   */
  static endTimer(label: string): number {
    const startTime = this.timers.get(label);
    if (!startTime) {
      throw new Error(`Timer not found: ${label}`);
    }

    const elapsed = Date.now() - startTime;
    this.timers.delete(label);
    return elapsed;
  }

  /**
   * Measure function execution time
   */
  static async measureAsync<T>(
    fn: () => Promise<T>,
    label?: string
  ): Promise<{ result: T; elapsed: number }> {
    const timerLabel = label || 'anonymous';
    this.startTimer(timerLabel);
    
    try {
      const result = await fn();
      const elapsed = this.endTimer(timerLabel);
      return { result, elapsed };
    } catch (error) {
      this.timers.delete(timerLabel);
      throw error;
    }
  }

  /**
   * Memory usage snapshot
   */
  static getMemoryUsage(): NodeJS.MemoryUsage {
    return process.memoryUsage();
  }

  /**
   * Format bytes to human readable string
   */
  static formatBytes(bytes: number): string {
    const units = ['B', 'KB', 'MB', 'GB', 'TB'];
    let size = bytes;
    let unitIndex = 0;

    while (size >= 1024 && unitIndex < units.length - 1) {
      size /= 1024;
      unitIndex++;
    }

    return `${size.toFixed(2)} ${units[unitIndex]}`;
  }
}

/**
 * Async utilities
 */
export class AsyncUtils {
  /**
   * Delay execution
   */
  static delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Retry operation with exponential backoff
   */
  static async retry<T>(
    operation: () => Promise<T>,
    maxAttempts: number = 3,
    baseDelay: number = 1000
  ): Promise<T> {
    let lastError: Error;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      try {
        return await operation();
      } catch (error) {
        lastError = error as Error;
        
        if (attempt === maxAttempts) {
          break;
        }

        const delay = baseDelay * Math.pow(2, attempt - 1);
        logger.warn(`Operation failed, retrying in ${delay}ms (attempt ${attempt}/${maxAttempts})`, {
          error: error instanceof Error ? error.message : String(error),
        });
        
        await this.delay(delay);
      }
    }

    throw lastError!;
  }

  /**
   * Process items in batches
   */
  static async processBatch<T, R>(
    items: T[],
    processor: (item: T) => Promise<R>,
    batchSize: number = 10,
    onProgress?: (completed: number, total: number) => void
  ): Promise<R[]> {
    const results: R[] = [];
    
    for (let i = 0; i < items.length; i += batchSize) {
      const batch = items.slice(i, i + batchSize);
      const batchResults = await Promise.all(batch.map(processor));
      results.push(...batchResults);
      
      if (onProgress) {
        onProgress(Math.min(i + batchSize, items.length), items.length);
      }
    }

    return results;
  }
}
