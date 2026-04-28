/**
 * Rate Limiting Middleware
 * Provides per-route rate limiting for different API endpoints
 */

import rateLimit, { RateLimitRequestHandler } from 'express-rate-limit';
import { env } from '../config/env';
import { ApiError, ErrorCode } from './errorHandler';

// Rate limit response handler
const rateLimitHandler = (message: string) => {
    throw ApiError.rateLimited(message);
};

/**
 * General API rate limiter
 * 100 requests per 15 minutes
 */
export const apiLimiter: RateLimitRequestHandler = rateLimit({
    windowMs: env.RATE_LIMIT_WINDOW_MS,
    max: env.RATE_LIMIT_MAX_REQUESTS,
    message: 'Too many requests from this IP, please try again later',
    standardHeaders: true,
    legacyHeaders: false,
    handler: (req, res) => {
        res.status(429).json({
            success: false,
            error: {
                message: 'Too many requests from this IP, please try again later',
                code: ErrorCode.RATE_LIMITED,
                type: 'server',
                timestamp: new Date().toISOString(),
            },
        });
    },
});

/**
 * Strict rate limiter for authentication endpoints
 * 5 attempts per 15 minutes
 */
export const authLimiter: RateLimitRequestHandler = rateLimit({
    windowMs: 15 * 60 * 1000, // 15 minutes
    max: 5, // 5 attempts
    message: 'Too many authentication attempts, please try again later',
    standardHeaders: true,
    legacyHeaders: false,
    skipSuccessfulRequests: false,
    handler: (req, res) => {
        res.status(429).json({
            success: false,
            error: {
                message: 'Too many authentication attempts, please try again later',
                code: ErrorCode.RATE_LIMITED,
                type: 'authentication',
                timestamp: new Date().toISOString(),
            },
        });
    },
});

/**
 * Email sending rate limiter
 * 10 email campaigns per hour
 */
export const emailLimiter: RateLimitRequestHandler = rateLimit({
    windowMs: 60 * 60 * 1000, // 1 hour
    max: 10, // 10 campaigns per hour
    message: 'Email sending limit reached, please try again later',
    standardHeaders: true,
    legacyHeaders: false,
    keyGenerator: (req) => {
        // Rate limit by user ID instead of IP for authenticated endpoints
        return (req as any).user?.id || req.ip || 'unknown';
    },
    handler: (req, res) => {
        res.status(429).json({
            success: false,
            error: {
                message: 'Email sending limit reached. You can send 10 campaigns per hour.',
                code: ErrorCode.EMAIL_RATE_LIMITED,
                type: 'email',
                timestamp: new Date().toISOString(),
            },
        });
    },
});

/**
 * File upload rate limiter
 * 20 uploads per hour
 */
export const uploadLimiter: RateLimitRequestHandler = rateLimit({
    windowMs: 60 * 60 * 1000, // 1 hour
    max: 20, // 20 uploads per hour
    message: 'Upload limit reached, please try again later',
    standardHeaders: true,
    legacyHeaders: false,
    handler: (req, res) => {
        res.status(429).json({
            success: false,
            error: {
                message: 'Upload limit reached, please try again later',
                code: ErrorCode.RATE_LIMITED,
                type: 'file',
                timestamp: new Date().toISOString(),
            },
        });
    },
});

/**
 * Credential creation rate limiter
 * 5 credentials per hour
 */
export const credentialLimiter: RateLimitRequestHandler = rateLimit({
    windowMs: 60 * 60 * 1000, // 1 hour
    max: 5, // 5 credentials per hour
    message: 'Credential creation limit reached',
    standardHeaders: true,
    legacyHeaders: false,
    handler: (req, res) => {
        res.status(429).json({
            success: false,
            error: {
                message: 'You can only add 5 email credentials per hour',
                code: ErrorCode.RATE_LIMITED,
                type: 'server',
                timestamp: new Date().toISOString(),
            },
        });
    },
});
