import dotenv from 'dotenv';
import {
    SECURITY_REQUIREMENTS,
    validateKeyLength,
    logSecurityWarnings,
} from '../utils/security';

dotenv.config();

interface EnvConfig {
    NODE_ENV: string;
    PORT: number;
    DATABASE_URL: string;
    JWT_SECRET: string;
    JWT_EXPIRES_IN: string;
    REFRESH_TOKEN_SECRET: string;
    REFRESH_TOKEN_EXPIRES_IN: string;
    ENCRYPTION_KEY: string;
    FRONTEND_URL: string;
    MAX_FILE_SIZE: number;
    UPLOAD_DIR: string;
    RATE_LIMIT_WINDOW_MS: number;
    RATE_LIMIT_MAX_REQUESTS: number;
}

const getEnvVar = (key: string, defaultValue?: string): string => {
    const value = process.env[key] || defaultValue;
    if (!value) {
        throw new Error(`Missing environment variable: ${key}`);
    }
    return value;
};

// Load environment variables
const NODE_ENV = getEnvVar('NODE_ENV', 'development');
const ENCRYPTION_KEY = getEnvVar('ENCRYPTION_KEY');
const JWT_SECRET = getEnvVar('JWT_SECRET');
const REFRESH_TOKEN_SECRET = getEnvVar('REFRESH_TOKEN_SECRET');

// Validate security keys on startup
validateKeyLength(
    ENCRYPTION_KEY,
    SECURITY_REQUIREMENTS.ENCRYPTION_KEY_LENGTH,
    'ENCRYPTION_KEY'
);

validateKeyLength(
    JWT_SECRET,
    SECURITY_REQUIREMENTS.JWT_SECRET_MIN_LENGTH,
    'JWT_SECRET'
);

validateKeyLength(
    REFRESH_TOKEN_SECRET,
    SECURITY_REQUIREMENTS.REFRESH_TOKEN_SECRET_MIN_LENGTH,
    'REFRESH_TOKEN_SECRET'
);

// Log security warnings (only once at startup)
logSecurityWarnings({
    nodeEnv: NODE_ENV,
    encryptionKey: ENCRYPTION_KEY,
    jwtSecret: JWT_SECRET,
    refreshTokenSecret: REFRESH_TOKEN_SECRET,
});

export const env: EnvConfig = {
    NODE_ENV,
    PORT: parseInt(getEnvVar('PORT', '5000'), 10),
    DATABASE_URL: getEnvVar('DATABASE_URL'),
    JWT_SECRET,
    JWT_EXPIRES_IN: getEnvVar('JWT_EXPIRES_IN', '15m'),
    REFRESH_TOKEN_SECRET,
    REFRESH_TOKEN_EXPIRES_IN: getEnvVar('REFRESH_TOKEN_EXPIRES_IN', '7d'),
    ENCRYPTION_KEY,
    FRONTEND_URL: getEnvVar('FRONTEND_URL', 'http://localhost:3000'),
    MAX_FILE_SIZE: parseInt(getEnvVar('MAX_FILE_SIZE', '10485760'), 10),
    UPLOAD_DIR: getEnvVar('UPLOAD_DIR', './uploads'),
    RATE_LIMIT_WINDOW_MS: parseInt(getEnvVar('RATE_LIMIT_WINDOW_MS', '900000'), 10),
    RATE_LIMIT_MAX_REQUESTS: parseInt(getEnvVar('RATE_LIMIT_MAX_REQUESTS', '100'), 10),
};

