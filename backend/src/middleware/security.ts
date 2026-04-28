/**
 * Security Headers Middleware
 * Configures CSP, CORS, and other security headers
 */

import { Request, Response, NextFunction } from 'express';
import helmet from 'helmet';
import cors from 'cors';
import { env } from '../config/env';

/**
 * Configure Helmet with API-friendly settings
 */
export const securityHeaders = helmet({
    contentSecurityPolicy: false, // Disable for API (frontend handles this)
    crossOriginEmbedderPolicy: false,
    crossOriginOpenerPolicy: false,
    crossOriginResourcePolicy: false, // Allow cross-origin requests
    dnsPrefetchControl: { allow: false },
    frameguard: { action: 'deny' },
    hidePoweredBy: true,
    hsts: env.NODE_ENV === 'production' ? {
        maxAge: 31536000,
        includeSubDomains: true,
        preload: true,
    } : false,
    ieNoOpen: true,
    noSniff: true,
    originAgentCluster: true,
    permittedCrossDomainPolicies: { permittedPolicies: 'none' },
    referrerPolicy: { policy: 'strict-origin-when-cross-origin' },
    xssFilter: true,
});

/**
 * CORS configuration - simplified for development
 */
export const corsConfig = cors({
    origin: env.NODE_ENV === 'development'
        ? true  // Allow all origins in development
        : env.FRONTEND_URL,
    credentials: true,
    methods: ['GET', 'POST', 'PUT', 'PATCH', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With', 'X-CSRF-Token'],
    exposedHeaders: ['X-RateLimit-Limit', 'X-RateLimit-Remaining', 'X-RateLimit-Reset'],
    maxAge: 86400,
});

/**
 * CSRF protection middleware
 * Uses double-submit cookie pattern for stateless CSRF protection
 */
export const csrfProtection = (req: Request, res: Response, next: NextFunction): void => {
    // Skip CSRF for GET, HEAD, OPTIONS requests
    if (['GET', 'HEAD', 'OPTIONS'].includes(req.method)) {
        next();
        return;
    }

    // Skip CSRF in development mode
    if (env.NODE_ENV === 'development') {
        next();
        return;
    }

    // Check for CSRF token in header
    const csrfToken = req.headers['x-csrf-token'] as string;
    const csrfCookie = req.cookies?.['csrf-token'];

    // If no CSRF setup yet, skip (first request)
    if (!csrfCookie && !csrfToken) {
        next();
        return;
    }

    // Validate CSRF token
    if (!csrfToken || csrfToken !== csrfCookie) {
        res.status(403).json({
            success: false,
            error: {
                message: 'Invalid CSRF token',
                code: 'AUTH_006',
                type: 'authentication',
                timestamp: new Date().toISOString(),
            },
        });
        return;
    }

    next();
};

/**
 * Generate CSRF token for client
 */
export const generateCsrfToken = (req: Request, res: Response): void => {
    const crypto = require('crypto');
    const token = crypto.randomBytes(32).toString('hex');

    res.cookie('csrf-token', token, {
        httpOnly: true,
        secure: env.NODE_ENV === 'production',
        sameSite: 'strict',
        maxAge: 3600000, // 1 hour
    });

    res.json({ csrfToken: token });
};

/**
 * Additional security headers for API responses
 */
export const apiSecurityHeaders = (req: Request, res: Response, next: NextFunction): void => {
    // Prevent caching of sensitive data
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');

    // Additional security headers
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('X-Frame-Options', 'DENY');
    res.setHeader('X-XSS-Protection', '1; mode=block');

    next();
};
