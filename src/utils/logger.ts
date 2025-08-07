import chalk from 'chalk';
import { LoggerOptions } from '../types';

/**
 * Enhanced logger with structured logging and performance tracking
 */
export class Logger {
  private options: LoggerOptions;
  private startTime: number = Date.now();

  constructor(options: Partial<LoggerOptions> = {}) {
    this.options = {
      level: 'info',
      timestamp: true,
      colors: true,
      stream: process.stdout,
      ...options,
    };
  }

  private shouldLog(level: string): boolean {
    const levels = ['debug', 'info', 'warn', 'error'];
    const currentLevelIndex = levels.indexOf(this.options.level);
    const messageLevelIndex = levels.indexOf(level);
    return messageLevelIndex >= currentLevelIndex;
  }

  private formatMessage(level: string, message: string, meta?: Record<string, unknown>): string {
    let formatted = '';

    if (this.options.timestamp) {
      const timestamp = new Date().toISOString();
      formatted += this.options.colors ? chalk.gray(`[${timestamp}]`) : `[${timestamp}]`;
      formatted += ' ';
    }

    const levelFormatted = this.options.colors ? this.colorizeLevel(level) : level.toUpperCase();
    formatted += `${levelFormatted}: ${message}`;

    if (meta && Object.keys(meta).length > 0) {
      formatted += ` ${JSON.stringify(meta)}`;
    }

    return formatted;
  }

  private colorizeLevel(level: string): string {
    switch (level) {
      case 'debug':
        return chalk.blue('DEBUG');
      case 'info':
        return chalk.green('INFO');
      case 'warn':
        return chalk.yellow('WARN');
      case 'error':
        return chalk.red('ERROR');
      default:
        return level.toUpperCase();
    }
  }

  debug(message: string, meta?: Record<string, unknown>): void {
    if (this.shouldLog('debug')) {
      const formatted = this.formatMessage('debug', message, meta);
      this.options.stream?.write(formatted + '\n');
    }
  }

  info(message: string, meta?: Record<string, unknown>): void {
    if (this.shouldLog('info')) {
      const formatted = this.formatMessage('info', message, meta);
      this.options.stream?.write(formatted + '\n');
    }
  }

  warn(message: string, meta?: Record<string, unknown>): void {
    if (this.shouldLog('warn')) {
      const formatted = this.formatMessage('warn', message, meta);
      this.options.stream?.write(formatted + '\n');
    }
  }

  error(message: string, meta?: Record<string, unknown>): void {
    if (this.shouldLog('error')) {
      const formatted = this.formatMessage('error', message, meta);
      (this.options.stream as NodeJS.WriteStream || process.stderr).write(formatted + '\n');
    }
  }

  time(label: string): void {
    this.debug(`Timer started: ${label}`);
  }

  timeEnd(label: string): void {
    const elapsed = Date.now() - this.startTime;
    this.debug(`Timer ended: ${label} (${elapsed}ms)`);
  }

  group(label: string): void {
    this.info(`--- ${label} ---`);
  }

  groupEnd(): void {
    this.info('--- End ---');
  }
}

/**
 * Global logger instance
 */
export const logger = new Logger();
