/**
 * Enhanced Error Handling System
 * Provides structured errors with codes, types, and detailed context
 */

import { Request, Response, NextFunction } from 'express';
import { ZodError } from 'zod';

// Error codes for API responses
export enum ErrorCode {
    // Authentication errors (1xxx)
    INVALID_CREDENTIALS = 'AUTH_001',
    TOKEN_EXPIRED = 'AUTH_002',
    TOKEN_INVALID = 'AUTH_003',
    UNAUTHORIZED = 'AUTH_004',
    REFRESH_TOKEN_INVALID = 'AUTH_005',

    // Validation errors (2xxx)
    VALIDATION_FAILED = 'VAL_001',
    INVALID_INPUT = 'VAL_002',
    MISSING_REQUIRED_FIELD = 'VAL_003',
    INVALID_FORMAT = 'VAL_004',

    // Resource errors (3xxx)
    RESOURCE_NOT_FOUND = 'RES_001',
    RESOURCE_ALREADY_EXISTS = 'RES_002',
    RESOURCE_CONFLICT = 'RES_003',

    // Email errors (4xxx)
    EMAIL_SEND_FAILED = 'EMAIL_001',
    INVALID_SMTP_CREDENTIALS = 'EMAIL_002',
    EMAIL_RATE_LIMITED = 'EMAIL_003',

    // File errors (5xxx)
    FILE_TOO_LARGE = 'FILE_001',
    INVALID_FILE_TYPE = 'FILE_002',
    FILE_UPLOAD_FAILED = 'FILE_003',

    // Server errors (9xxx)
    INTERNAL_ERROR = 'SRV_001',
    DATABASE_ERROR = 'SRV_002',
    ENCRYPTION_ERROR = 'SRV_003',
    RATE_LIMITED = 'SRV_004',
}

// Error types for categorization
export type ErrorType =
    | 'authentication'
    | 'validation'
    | 'resource'
    | 'email'
    | 'file'
    | 'server';

/**
 * Enhanced API Error class with error codes and types
 */
export class ApiError extends Error {
    readonly statusCode: number;
    readonly code: ErrorCode;
    readonly type: ErrorType;
    readonly isOperational: boolean;
    readonly details?: Record<string, unknown>;
    readonly timestamp: string;

    constructor(
        message: string,
        statusCode: number,
        code: ErrorCode,
        type: ErrorType,
        details?: Record<string, unknown>
    ) {
        super(message);
        this.name = 'ApiError';
        this.statusCode = statusCode;
        this.code = code;
        this.type = type;
        this.isOperational = true;
        this.details = details;
        this.timestamp = new Date().toISOString();
        Error.captureStackTrace(this, this.constructor);
    }

    // Factory methods for common errors
    static badRequest(message: string, code = ErrorCode.INVALID_INPUT, details?: Record<string, unknown>) {
        return new ApiError(message, 400, code, 'validation', details);
    }

    static unauthorized(message = 'Unauthorized', code = ErrorCode.UNAUTHORIZED) {
        return new ApiError(message, 401, code, 'authentication');
    }

    static forbidden(message = 'Access denied') {
        return new ApiError(message, 403, ErrorCode.UNAUTHORIZED, 'authentication');
    }

    static notFound(resource = 'Resource') {
        return new ApiError(`${resource} not found`, 404, ErrorCode.RESOURCE_NOT_FOUND, 'resource');
    }

    static conflict(message: string, code = ErrorCode.RESOURCE_CONFLICT) {
        return new ApiError(message, 409, code, 'resource');
    }

    static validationFailed(errors: Array<{ field: string; message: string }>) {
        return new ApiError(
            'Validation failed',
            400,
            ErrorCode.VALIDATION_FAILED,
            'validation',
            { errors }
        );
    }

    static fromZodError(error: ZodError): ApiError {
        const errors = error.errors.map(err => ({
            field: err.path.join('.'),
            message: err.message,
        }));
        return ApiError.validationFailed(errors);
    }

    static internal(message = 'Internal server error') {
        return new ApiError(message, 500, ErrorCode.INTERNAL_ERROR, 'server');
    }

    static emailError(message: string, code = ErrorCode.EMAIL_SEND_FAILED) {
        return new ApiError(message, 500, code, 'email');
    }

    static fileError(message: string, code = ErrorCode.FILE_UPLOAD_FAILED) {
        return new ApiError(message, 400, code, 'file');
    }

    static rateLimited(message = 'Too many requests') {
        return new ApiError(message, 429, ErrorCode.RATE_LIMITED, 'server');
    }

    toJSON() {
        return {
            success: false,
            error: {
                message: this.message,
                code: this.code,
                type: this.type,
                ...(this.details && { details: this.details }),
                timestamp: this.timestamp,
            },
        };
    }
}

// Legacy AppError for backward compatibility
export class AppError extends ApiError {
    constructor(message: string, statusCode: number = 500) {
        super(
            message,
            statusCode,
            statusCode === 400 ? ErrorCode.INVALID_INPUT :
                statusCode === 401 ? ErrorCode.UNAUTHORIZED :
                    statusCode === 404 ? ErrorCode.RESOURCE_NOT_FOUND :
                        ErrorCode.INTERNAL_ERROR,
            statusCode === 400 ? 'validation' :
                statusCode === 401 ? 'authentication' :
                    statusCode === 404 ? 'resource' :
                        'server'
        );
    }
}

/**
 * Global error handler middleware
 */
export const errorHandler = (
    err: Error,
    req: Request,
    res: Response,
    _next: NextFunction
): void => {
    // Handle Zod validation errors
    if (err instanceof ZodError) {
        const apiError = ApiError.fromZodError(err);
        res.status(apiError.statusCode).json(apiError.toJSON());
        return;
    }

    // Handle our custom API errors
    if (err instanceof ApiError) {
        res.status(err.statusCode).json(err.toJSON());
        return;
    }

    // Handle multer errors
    if (err.name === 'MulterError') {
        const apiError = ApiError.fileError(err.message, ErrorCode.FILE_UPLOAD_FAILED);
        res.status(apiError.statusCode).json(apiError.toJSON());
        return;
    }

    // Log unexpected errors
    console.error('Unexpected error:', {
        message: err.message,
        stack: err.stack,
        path: req.path,
        method: req.method,
    });

    // Generic error response
    const apiError = ApiError.internal(
        process.env.NODE_ENV === 'development' ? err.message : 'Internal server error'
    );

    res.status(apiError.statusCode).json({
        ...apiError.toJSON(),
        ...(process.env.NODE_ENV === 'development' && { stack: err.stack }),
    });
};

/**
 * 404 Not Found handler
 */
export const notFound = (req: Request, res: Response): void => {
    const error = ApiError.notFound('Route');
    res.status(error.statusCode).json(error.toJSON());
};

/**
 * Async handler wrapper to catch promise rejections
 */
export const asyncHandler = (
    fn: (req: Request, res: Response, next: NextFunction) => Promise<void>
) => {
    return (req: Request, res: Response, next: NextFunction) => {
        Promise.resolve(fn(req, res, next)).catch(next);
    };
};
