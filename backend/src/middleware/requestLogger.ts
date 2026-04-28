/**
 * Request ID Middleware
 * Adds a unique request ID to each request for tracing
 */
import { Request, Response, NextFunction } from 'express';
import { v4 as uuidv4 } from 'uuid';
import logger from '../utils/logger';

// Extend Express Request type
declare global {
    namespace Express {
        interface Request {
            requestId: string;
            startTime: number;
        }
    }
}

/**
 * Adds a unique request ID and start time to each request
 */
export const requestIdMiddleware = (req: Request, res: Response, next: NextFunction) => {
    // Generate or use existing request ID
    req.requestId = (req.headers['x-request-id'] as string) || uuidv4();
    req.startTime = Date.now();

    // Add request ID to response headers
    res.setHeader('X-Request-ID', req.requestId);

    next();
};

/**
 * Logs request details after response is sent
 */
export const requestLoggerMiddleware = (req: Request, res: Response, next: NextFunction) => {
    // Log after response
    res.on('finish', () => {
        const duration = Date.now() - req.startTime;
        const logData = {
            requestId: req.requestId,
            method: req.method,
            path: req.path,
            statusCode: res.statusCode,
            duration: `${duration}ms`,
            userAgent: req.get('User-Agent')?.substring(0, 50),
        };

        if (res.statusCode >= 500) {
            logger.error('Request failed', logData);
        } else if (res.statusCode >= 400) {
            logger.warn('Request error', logData);
        } else {
            logger.http('Request completed', logData);
        }
    });

    next();
};
