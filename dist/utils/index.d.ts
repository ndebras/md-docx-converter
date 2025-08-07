import { ValidationResult } from '../types';
/**
 * File system utilities with enhanced error handling
 */
export declare class FileUtils {
    /**
     * Safely read file with encoding detection
     */
    static readFile(filePath: string): Promise<string>;
    /**
     * Safely write file with directory creation
     */
    static writeFile(filePath: string, content: string | Buffer): Promise<void>;
    /**
     * Get file extension
     */
    static getExtension(filePath: string): string;
    /**
     * Get MIME type from file extension
     */
    static getMimeType(filePath: string): string | false;
    /**
     * Validate file path and type
     */
    static validateFile(filePath: string, allowedExtensions: string[]): ValidationResult;
    /**
     * Generate unique filename
     */
    static generateUniqueFileName(originalName: string, suffix?: string): string;
    /**
     * Sanitize filename for cross-platform compatibility
     */
    static sanitizeFileName(fileName: string): string;
    /**
     * Get relative path for output
     */
    static getRelativePath(from: string, to: string): string;
    /**
     * Ensure directory exists and is writable
     */
    static ensureWritableDir(dirPath: string): Promise<void>;
}
/**
 * String utilities for text processing
 */
export declare class StringUtils {
    /**
     * Generate URL-safe slug from text
     */
    static slugify(text: string): string;
    /**
     * Extract heading anchor from text
     */
    static createAnchor(heading: string): string;
    /**
     * Escape HTML special characters
     */
    static escapeHtml(text: string): string;
    /**
     * Unescape HTML entities
     */
    static unescapeHtml(text: string): string;
    /**
     * Clean whitespace and normalize line endings
     */
    static normalizeWhitespace(text: string): string;
    /**
     * Word wrap text to specified width
     */
    static wordWrap(text: string, width?: number): string;
    /**
     * Extract text content from HTML
     */
    static stripHtml(html: string): string;
    /**
     * Count words in text
     */
    static countWords(text: string): number;
    /**
     * Estimate reading time in minutes
     */
    static estimateReadingTime(text: string, wordsPerMinute?: number): number;
}
/**
 * Performance utilities
 */
export declare class PerformanceUtils {
    private static timers;
    /**
     * Start performance timer
     */
    static startTimer(label: string): void;
    /**
     * End performance timer and return elapsed time
     */
    static endTimer(label: string): number;
    /**
     * Measure function execution time
     */
    static measureAsync<T>(fn: () => Promise<T>, label?: string): Promise<{
        result: T;
        elapsed: number;
    }>;
    /**
     * Memory usage snapshot
     */
    static getMemoryUsage(): NodeJS.MemoryUsage;
    /**
     * Format bytes to human readable string
     */
    static formatBytes(bytes: number): string;
}
/**
 * Async utilities
 */
export declare class AsyncUtils {
    /**
     * Delay execution
     */
    static delay(ms: number): Promise<void>;
    /**
     * Retry operation with exponential backoff
     */
    static retry<T>(operation: () => Promise<T>, maxAttempts?: number, baseDelay?: number): Promise<T>;
    /**
     * Process items in batches
     */
    static processBatch<T, R>(items: T[], processor: (item: T) => Promise<R>, batchSize?: number, onProgress?: (completed: number, total: number) => void): Promise<R[]>;
}
//# sourceMappingURL=index.d.ts.map