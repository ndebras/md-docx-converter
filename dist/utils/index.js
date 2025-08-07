"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.AsyncUtils = exports.PerformanceUtils = exports.StringUtils = exports.FileUtils = void 0;
const fs = __importStar(require("fs-extra"));
const path = __importStar(require("path"));
const uuid_1 = require("uuid");
const mime = __importStar(require("mime-types"));
const logger_1 = require("./logger");
/**
 * File system utilities with enhanced error handling
 */
class FileUtils {
    /**
     * Safely read file with encoding detection
     */
    static async readFile(filePath) {
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
        }
        catch (error) {
            logger_1.logger.error(`Failed to read file: ${filePath}`, { error });
            throw error;
        }
    }
    /**
     * Safely write file with directory creation
     */
    static async writeFile(filePath, content) {
        try {
            const dir = path.dirname(filePath);
            await fs.ensureDir(dir);
            await fs.writeFile(filePath, content);
            logger_1.logger.debug(`File written successfully: ${filePath}`);
        }
        catch (error) {
            logger_1.logger.error(`Failed to write file: ${filePath}`, { error });
            throw error;
        }
    }
    /**
     * Get file extension
     */
    static getExtension(filePath) {
        return path.extname(filePath).toLowerCase();
    }
    /**
     * Get MIME type from file extension
     */
    static getMimeType(filePath) {
        return mime.lookup(filePath);
    }
    /**
     * Validate file path and type
     */
    static validateFile(filePath, allowedExtensions) {
        const errors = [];
        const warnings = [];
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
    static generateUniqueFileName(originalName, suffix) {
        const ext = path.extname(originalName);
        const baseName = path.basename(originalName, ext);
        const uniqueId = (0, uuid_1.v4)().slice(0, 8);
        const suffixPart = suffix ? `_${suffix}` : '';
        return `${baseName}${suffixPart}_${uniqueId}${ext}`;
    }
    /**
     * Sanitize filename for cross-platform compatibility
     */
    static sanitizeFileName(fileName) {
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
    static getRelativePath(from, to) {
        return path.relative(from, to).replace(/\\/g, '/');
    }
    /**
     * Ensure directory exists and is writable
     */
    static async ensureWritableDir(dirPath) {
        try {
            await fs.ensureDir(dirPath);
            // Test write permissions
            const testFile = path.join(dirPath, '.write_test');
            await fs.writeFile(testFile, 'test');
            await fs.remove(testFile);
            logger_1.logger.debug(`Directory ready: ${dirPath}`);
        }
        catch (error) {
            logger_1.logger.error(`Failed to ensure writable directory: ${dirPath}`, { error });
            throw new Error(`Cannot write to directory: ${dirPath}`);
        }
    }
}
exports.FileUtils = FileUtils;
/**
 * String utilities for text processing
 */
class StringUtils {
    /**
     * Generate URL-safe slug from text
     */
    static slugify(text) {
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
    static createAnchor(heading) {
        return this.slugify(heading);
    }
    /**
     * Escape HTML special characters
     */
    static escapeHtml(text) {
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
    static unescapeHtml(text) {
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
    static normalizeWhitespace(text) {
        return text
            .replace(/\r\n/g, '\n')
            .replace(/\r/g, '\n')
            .replace(/[ \t]+$/gm, '')
            .replace(/\n{3,}/g, '\n\n');
    }
    /**
     * Word wrap text to specified width
     */
    static wordWrap(text, width = 80) {
        const lines = text.split('\n');
        const wrappedLines = [];
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
                }
                else {
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
    static stripHtml(html) {
        return html.replace(/<[^>]*>/g, '').trim();
    }
    /**
     * Count words in text
     */
    static countWords(text) {
        return text.trim().split(/\s+/).filter(word => word.length > 0).length;
    }
    /**
     * Estimate reading time in minutes
     */
    static estimateReadingTime(text, wordsPerMinute = 200) {
        const wordCount = this.countWords(text);
        return Math.ceil(wordCount / wordsPerMinute);
    }
}
exports.StringUtils = StringUtils;
/**
 * Performance utilities
 */
class PerformanceUtils {
    /**
     * Start performance timer
     */
    static startTimer(label) {
        this.timers.set(label, Date.now());
    }
    /**
     * End performance timer and return elapsed time
     */
    static endTimer(label) {
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
    static async measureAsync(fn, label) {
        const timerLabel = label || 'anonymous';
        this.startTimer(timerLabel);
        try {
            const result = await fn();
            const elapsed = this.endTimer(timerLabel);
            return { result, elapsed };
        }
        catch (error) {
            this.timers.delete(timerLabel);
            throw error;
        }
    }
    /**
     * Memory usage snapshot
     */
    static getMemoryUsage() {
        return process.memoryUsage();
    }
    /**
     * Format bytes to human readable string
     */
    static formatBytes(bytes) {
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
exports.PerformanceUtils = PerformanceUtils;
PerformanceUtils.timers = new Map();
/**
 * Async utilities
 */
class AsyncUtils {
    /**
     * Delay execution
     */
    static delay(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }
    /**
     * Retry operation with exponential backoff
     */
    static async retry(operation, maxAttempts = 3, baseDelay = 1000) {
        let lastError;
        for (let attempt = 1; attempt <= maxAttempts; attempt++) {
            try {
                return await operation();
            }
            catch (error) {
                lastError = error;
                if (attempt === maxAttempts) {
                    break;
                }
                const delay = baseDelay * Math.pow(2, attempt - 1);
                logger_1.logger.warn(`Operation failed, retrying in ${delay}ms (attempt ${attempt}/${maxAttempts})`, {
                    error: error instanceof Error ? error.message : String(error),
                });
                await this.delay(delay);
            }
        }
        throw lastError;
    }
    /**
     * Process items in batches
     */
    static async processBatch(items, processor, batchSize = 10, onProgress) {
        const results = [];
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
exports.AsyncUtils = AsyncUtils;
//# sourceMappingURL=index.js.map