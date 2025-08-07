import { LoggerOptions } from '../types';
/**
 * Enhanced logger with structured logging and performance tracking
 */
export declare class Logger {
    private options;
    private startTime;
    constructor(options?: Partial<LoggerOptions>);
    private shouldLog;
    private formatMessage;
    private colorizeLevel;
    debug(message: string, meta?: Record<string, unknown>): void;
    info(message: string, meta?: Record<string, unknown>): void;
    warn(message: string, meta?: Record<string, unknown>): void;
    error(message: string, meta?: Record<string, unknown>): void;
    time(label: string): void;
    timeEnd(label: string): void;
    group(label: string): void;
    groupEnd(): void;
}
/**
 * Global logger instance
 */
export declare const logger: Logger;
//# sourceMappingURL=logger.d.ts.map