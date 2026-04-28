/**
 * Winston Logger Configuration
 * Provides structured logging for the application
 */

import winston from 'winston';
import path from 'path';
import { env } from '../config/env';

// Define log format
const logFormat = winston.format.combine(
    winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
    winston.format.errors({ stack: true }),
    winston.format.printf(({ timestamp, level, message, stack, ...meta }) => {
        let log = `${timestamp} [${level.toUpperCase()}]: ${message}`;
        if (Object.keys(meta).length > 0) {
            log += ` ${JSON.stringify(meta)}`;
        }
        if (stack) {
            log += `\n${stack}`;
        }
        return log;
    })
);

// Define log levels
const levels = {
    error: 0,
    warn: 1,
    info: 2,
    http: 3,
    debug: 4,
};

// Create transports based on environment
const transports: winston.transport[] = [
    // Console transport
    new winston.transports.Console({
        format: winston.format.combine(
            winston.format.colorize({ all: true }),
            logFormat
        ),
    }),
];

// Add file transports in production
if (env.NODE_ENV === 'production') {
    transports.push(
        // Error log file
        new winston.transports.File({
            filename: path.join('logs', 'error.log'),
            level: 'error',
            format: logFormat,
            maxsize: 5242880, // 5MB
            maxFiles: 5,
        }),
        // Combined log file
        new winston.transports.File({
            filename: path.join('logs', 'combined.log'),
            format: logFormat,
            maxsize: 5242880, // 5MB
            maxFiles: 5,
        })
    );
}

// Create logger instance
const logger = winston.createLogger({
    level: env.NODE_ENV === 'development' ? 'debug' : 'info',
    levels,
    format: logFormat,
    transports,
    exitOnError: false,
});

// Create stream for Morgan HTTP logging
export const httpLogStream = {
    write: (message: string) => {
        logger.http(message.trim());
    },
};

// Export logger with typed methods
export default logger;

// Convenience exports
export const logInfo = (message: string, meta?: object) => logger.info(message, meta);
export const logError = (message: string, error?: Error, meta?: object) => {
    logger.error(message, { error: error?.message, stack: error?.stack, ...meta });
};
export const logWarn = (message: string, meta?: object) => logger.warn(message, meta);
export const logDebug = (message: string, meta?: object) => logger.debug(message, meta);
export const logHttp = (message: string, meta?: object) => logger.http(message, meta);
