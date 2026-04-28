/**
 * Validation Middleware
 * Uses Zod schemas for request validation with enhanced error handling
 */

import { Request, Response, NextFunction } from 'express';
import { ZodSchema, ZodError } from 'zod';
import { ApiError } from './errorHandler';

/**
 * Validates request body against a Zod schema
 */
export const validate = (schema: ZodSchema) => {
    return (req: Request, res: Response, next: NextFunction): void => {
        try {
            req.body = schema.parse(req.body);
            next();
        } catch (error) {
            if (error instanceof ZodError) {
                next(ApiError.fromZodError(error));
                return;
            }
            next(ApiError.badRequest('Invalid request data'));
        }
    };
};

/**
 * Validates query parameters against a Zod schema
 */
export const validateQuery = (schema: ZodSchema) => {
    return (req: Request, res: Response, next: NextFunction): void => {
        try {
            req.query = schema.parse(req.query);
            next();
        } catch (error) {
            if (error instanceof ZodError) {
                next(ApiError.fromZodError(error));
                return;
            }
            next(ApiError.badRequest('Invalid query parameters'));
        }
    };
};

/**
 * Validates URL parameters against a Zod schema
 */
export const validateParams = (schema: ZodSchema) => {
    return (req: Request, res: Response, next: NextFunction): void => {
        try {
            req.params = schema.parse(req.params);
            next();
        } catch (error) {
            if (error instanceof ZodError) {
                next(ApiError.fromZodError(error));
                return;
            }
            next(ApiError.badRequest('Invalid URL parameters'));
        }
    };
};
