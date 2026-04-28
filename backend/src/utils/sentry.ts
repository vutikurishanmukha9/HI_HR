/**
 * Sentry Error Tracking Configuration
 * 
 * To enable Sentry:
 * 1. Create account at https://sentry.io
 * 2. Get DSN from project settings
 * 3. Add SENTRY_DSN to .env
 */
import * as Sentry from '@sentry/node';
import { env } from '../config/env';
import logger from './logger';

// Check if Sentry is configured
const SENTRY_DSN = process.env.SENTRY_DSN;

/**
 * Initialize Sentry error tracking
 */
export function initSentry(): void {
    if (!SENTRY_DSN) {
        logger.info('Sentry DSN not configured - error tracking disabled');
        return;
    }

    Sentry.init({
        dsn: SENTRY_DSN,
        environment: env.NODE_ENV,
        enabled: env.NODE_ENV === 'production',

        // Performance monitoring
        tracesSampleRate: 0.1, // 10% of transactions

        // Filter out non-error events
        beforeSend(event) {
            // Don't send events in development
            if (env.NODE_ENV === 'development') {
                return null;
            }
            return event;
        },

        // Ignore common non-critical errors
        ignoreErrors: [
            'Network request failed',
            'Failed to fetch',
            'AbortError',
        ],
    });

    logger.info('Sentry error tracking initialized');
}

/**
 * Capture an exception with additional context
 */
export function captureException(error: Error, context?: Record<string, any>): void {
    if (!SENTRY_DSN) {
        logger.error('Error (Sentry disabled):', { error: error.message, ...context });
        return;
    }

    Sentry.withScope((scope) => {
        if (context) {
            scope.setExtras(context);
        }
        Sentry.captureException(error);
    });
}

/**
 * Set user context for error tracking
 */
export function setUserContext(userId: string, email?: string): void {
    if (!SENTRY_DSN) return;

    Sentry.setUser({
        id: userId,
        email: email,
    });
}

/**
 * Clear user context (on logout)
 */
export function clearUserContext(): void {
    if (!SENTRY_DSN) return;
    Sentry.setUser(null);
}

export default {
    init: initSentry,
    captureException,
    setUserContext,
    clearUserContext,
};

