import crypto from 'crypto';

/**
 * Security utilities for input sanitization, token generation, and validation
 */

// Minimum required lengths for security keys
export const SECURITY_REQUIREMENTS = {
    ENCRYPTION_KEY_LENGTH: 32,
    JWT_SECRET_MIN_LENGTH: 32,
    REFRESH_TOKEN_SECRET_MIN_LENGTH: 32,
} as const;

/**
 * Generates a cryptographically secure random string
 * @param length - Length of the string in bytes (output will be hex, so 2x chars)
 */
export const generateSecureToken = (length: number = 32): string => {
    return crypto.randomBytes(length).toString('hex');
};

/**
 * Generates a secure encryption key (32 bytes for AES-256)
 */
export const generateEncryptionKey = (): string => {
    return crypto.randomBytes(SECURITY_REQUIREMENTS.ENCRYPTION_KEY_LENGTH).toString('hex');
};

/**
 * Validates that a key meets minimum security requirements
 */
export const validateKeyLength = (key: string, minLength: number, keyName: string): void => {
    if (!key || key.length < minLength) {
        throw new Error(
            `${keyName} must be at least ${minLength} characters. ` +
            `Current length: ${key?.length || 0}. ` +
            `Generate a secure key with: node -e "console.log(require('crypto').randomBytes(${minLength}).toString('hex'))"`
        );
    }
};

/**
 * Checks if a secret appears to be a default/insecure value
 */
export const isInsecureSecret = (secret: string): boolean => {
    const insecurePatterns = [
        /^your-/i,
        /^change-this/i,
        /^secret$/i,
        /^password$/i,
        /^default/i,
        /^test/i,
        /^dev/i,
        /streammail/i,
        /example/i,
    ];

    return insecurePatterns.some(pattern => pattern.test(secret));
};

/**
 * Logs security warnings for development mode or insecure configurations
 */
export const logSecurityWarnings = (config: {
    nodeEnv: string;
    encryptionKey: string;
    jwtSecret: string;
    refreshTokenSecret: string;
}): void => {
    const warnings: string[] = [];

    if (config.nodeEnv === 'development') {
        console.warn('\n⚠️  Running in DEVELOPMENT mode - not suitable for production!\n');
    }

    if (isInsecureSecret(config.encryptionKey)) {
        warnings.push('ENCRYPTION_KEY appears to be a default/insecure value');
    }

    if (isInsecureSecret(config.jwtSecret)) {
        warnings.push('JWT_SECRET appears to be a default/insecure value');
    }

    if (isInsecureSecret(config.refreshTokenSecret)) {
        warnings.push('REFRESH_TOKEN_SECRET appears to be a default/insecure value');
    }

    if (warnings.length > 0) {
        console.warn('\n🔒 SECURITY WARNINGS:');
        warnings.forEach(w => console.warn(`   ⚠️  ${w}`));
        console.warn('\n   Generate secure keys with:');
        console.warn('   node -e "console.log(require(\'crypto\').randomBytes(32).toString(\'hex\'))"');
        console.warn('');
    }
};

/**
 * Sanitizes a string to prevent XSS attacks
 * Escapes HTML special characters
 */
export const sanitizeHtml = (input: string): string => {
    const htmlEscapes: Record<string, string> = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#x27;',
        '/': '&#x2F;',
    };

    return input.replace(/[&<>"'/]/g, char => htmlEscapes[char]);
};

/**
 * Sanitizes an object's string values recursively
 */
export const sanitizeObject = <T extends Record<string, any>>(obj: T): T => {
    const sanitized: Record<string, any> = {};

    for (const [key, value] of Object.entries(obj)) {
        if (typeof value === 'string') {
            sanitized[key] = sanitizeHtml(value);
        } else if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
            sanitized[key] = sanitizeObject(value);
        } else if (Array.isArray(value)) {
            sanitized[key] = value.map(item =>
                typeof item === 'string' ? sanitizeHtml(item) :
                    typeof item === 'object' && item !== null ? sanitizeObject(item) : item
            );
        } else {
            sanitized[key] = value;
        }
    }

    return sanitized as T;
};

/**
 * Validates and sanitizes email address
 */
export const sanitizeEmail = (email: string): string => {
    // Remove any whitespace and convert to lowercase
    const cleaned = email.trim().toLowerCase();

    // Basic email validation regex
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

    if (!emailRegex.test(cleaned)) {
        throw new Error('Invalid email format');
    }

    return cleaned;
};

/**
 * Strips potentially dangerous characters from file names
 */
export const sanitizeFileName = (fileName: string): string => {
    // Remove path traversal attempts and dangerous characters
    return fileName
        .replace(/\.\./g, '')
        .replace(/[<>:"/\\|?*\x00-\x1f]/g, '')
        .trim();
};
