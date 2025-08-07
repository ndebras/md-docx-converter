"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.logger = exports.Logger = void 0;
const chalk_1 = __importDefault(require("chalk"));
/**
 * Enhanced logger with structured logging and performance tracking
 */
class Logger {
    constructor(options = {}) {
        this.startTime = Date.now();
        this.options = {
            level: 'info',
            timestamp: true,
            colors: true,
            stream: process.stdout,
            ...options,
        };
    }
    shouldLog(level) {
        const levels = ['debug', 'info', 'warn', 'error'];
        const currentLevelIndex = levels.indexOf(this.options.level);
        const messageLevelIndex = levels.indexOf(level);
        return messageLevelIndex >= currentLevelIndex;
    }
    formatMessage(level, message, meta) {
        let formatted = '';
        if (this.options.timestamp) {
            const timestamp = new Date().toISOString();
            formatted += this.options.colors ? chalk_1.default.gray(`[${timestamp}]`) : `[${timestamp}]`;
            formatted += ' ';
        }
        const levelFormatted = this.options.colors ? this.colorizeLevel(level) : level.toUpperCase();
        formatted += `${levelFormatted}: ${message}`;
        if (meta && Object.keys(meta).length > 0) {
            formatted += ` ${JSON.stringify(meta)}`;
        }
        return formatted;
    }
    colorizeLevel(level) {
        switch (level) {
            case 'debug':
                return chalk_1.default.blue('DEBUG');
            case 'info':
                return chalk_1.default.green('INFO');
            case 'warn':
                return chalk_1.default.yellow('WARN');
            case 'error':
                return chalk_1.default.red('ERROR');
            default:
                return level.toUpperCase();
        }
    }
    debug(message, meta) {
        if (this.shouldLog('debug')) {
            const formatted = this.formatMessage('debug', message, meta);
            this.options.stream?.write(formatted + '\n');
        }
    }
    info(message, meta) {
        if (this.shouldLog('info')) {
            const formatted = this.formatMessage('info', message, meta);
            this.options.stream?.write(formatted + '\n');
        }
    }
    warn(message, meta) {
        if (this.shouldLog('warn')) {
            const formatted = this.formatMessage('warn', message, meta);
            this.options.stream?.write(formatted + '\n');
        }
    }
    error(message, meta) {
        if (this.shouldLog('error')) {
            const formatted = this.formatMessage('error', message, meta);
            (this.options.stream || process.stderr).write(formatted + '\n');
        }
    }
    time(label) {
        this.debug(`Timer started: ${label}`);
    }
    timeEnd(label) {
        const elapsed = Date.now() - this.startTime;
        this.debug(`Timer ended: ${label} (${elapsed}ms)`);
    }
    group(label) {
        this.info(`--- ${label} ---`);
    }
    groupEnd() {
        this.info('--- End ---');
    }
}
exports.Logger = Logger;
/**
 * Global logger instance
 */
exports.logger = new Logger();
//# sourceMappingURL=logger.js.map