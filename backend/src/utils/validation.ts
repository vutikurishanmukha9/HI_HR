import { z } from 'zod';

export const registerSchema = z.object({
    email: z.string().email('Invalid email address'),
    password: z
        .string()
        .min(8, 'Password must be at least 8 characters')
        .regex(/[A-Z]/, 'Password must contain at least one uppercase letter')
        .regex(/[a-z]/, 'Password must contain at least one lowercase letter')
        .regex(/[0-9]/, 'Password must contain at least one number'),
    fullName: z.string().min(2, 'Full name must be at least 2 characters'),
});

export const loginSchema = z.object({
    email: z.string().email('Invalid email address'),
    password: z.string().min(1, 'Password is required'),
});

export const emailCredentialSchema = z.object({
    email: z.string().email('Invalid email address'),
    appPassword: z.string().min(16, 'App password must be at least 16 characters'),
    isDefault: z.boolean().optional(),
});

export const createCampaignSchema = z.object({
    name: z.string().min(1, 'Campaign name is required'),
    subject: z.string().min(1, 'Subject is required').max(500, 'Subject too long'),
    body: z.string().min(1, 'Email body is required'),
    scheduledTime: z.string().datetime().optional(),
    batchSize: z.number().int().min(1).max(100).optional(),
    batchDelay: z.number().int().min(0).max(3600).optional(),
});

export const recipientSchema = z.object({
    fullName: z.string().min(1, 'Full name is required'),
    email: z.string().email('Invalid email address'),
    companyName: z.string().min(1, 'Company name is required'),
    jobTitle: z.string().optional(),
});

export const updateCampaignSchema = z.object({
    name: z.string().min(1).optional(),
    status: z.enum(['draft', 'scheduled', 'sending', 'completed', 'paused', 'failed']).optional(),
    scheduledTime: z.string().datetime().optional(),
    batchSize: z.number().int().min(1).max(100).optional(),
    batchDelay: z.number().int().min(0).max(3600).optional(),
});

export const paginationSchema = z.object({
    page: z.string().transform(Number).pipe(z.number().int().min(1)).optional(),
    limit: z.string().transform(Number).pipe(z.number().int().min(1).max(100)).optional(),
});

// Schema for sending campaign emails via backend
export const sendCampaignSchema = z.object({
    credentialEmail: z.string().email('Invalid sender email'),
    subject: z.string().min(1, 'Subject is required').max(500, 'Subject too long'),
    body: z.string().min(1, 'Email body is required'),
    recipients: z.array(z.object({
        email: z.string().email('Invalid recipient email'),
        fullName: z.string().min(1, 'Full name is required'),
        companyName: z.string().min(1, 'Company name is required'),
        jobTitle: z.string().optional(),
    })).min(1, 'At least one recipient is required').max(500, 'Maximum 500 recipients per batch'),
    attachments: z.array(z.object({
        filename: z.string(),
        content: z.string(), // base64 encoded
        contentType: z.string(),
    })).optional(),
    batchSize: z.number().int().min(1).max(50).default(10),
    batchDelay: z.number().int().min(0).max(300).default(60),
});

// Schema for refresh token
export const refreshTokenSchema = z.object({
    refreshToken: z.string().min(1, 'Refresh token is required'),
});

