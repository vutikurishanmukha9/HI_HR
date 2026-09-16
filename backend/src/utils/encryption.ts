import crypto from 'crypto';
import { env } from '../config/env';

const ALGORITHM = 'aes-256-cbc';
const IV_LENGTH = 16;

/**
 * Derives a full 32-byte key buffer using SHA-256 to ensure maximum entropy for AES-256
 */
const getKeyBuffer = (keyString: string): Buffer => {
    return crypto.createHash('sha256').update(keyString).digest();
};

/**
 * Encrypts sensitive data using AES-256-CBC
 */
export const encrypt = (text: string): string => {
    const iv = crypto.randomBytes(IV_LENGTH);
    const key = getKeyBuffer(env.ENCRYPTION_KEY);
    const cipher = crypto.createCipheriv(ALGORITHM, key, iv);

    let encrypted = cipher.update(text, 'utf8', 'hex');
    encrypted += cipher.final('hex');

    return iv.toString('hex') + ':' + encrypted;
};

/**
 * Decrypts data encrypted with the encrypt function.
 * Supports primary SHA-256 key derivation with legacy slice fallback.
 */
export const decrypt = (text: string): string => {
    const parts = text.split(':');
    if (parts.length < 2) {
        throw new Error('Invalid encrypted data format');
    }

    const iv = Buffer.from(parts.shift()!, 'hex');
    const encryptedText = parts.join(':');

    // Primary: SHA-256 derived key (256-bit entropy)
    try {
        const key = getKeyBuffer(env.ENCRYPTION_KEY);
        const decipher = crypto.createDecipheriv(ALGORITHM, key, iv);
        let decrypted = decipher.update(encryptedText, 'hex', 'utf8');
        decrypted += decipher.final('utf8');
        return decrypted;
    } catch {
        // Fallback: Legacy slice(0, 32) utf-8 derivation
        try {
            const legacyKey = Buffer.from(env.ENCRYPTION_KEY, 'utf-8').slice(0, 32);
            const legacyDecipher = crypto.createDecipheriv(ALGORITHM, legacyKey, iv);
            let decrypted = legacyDecipher.update(encryptedText, 'hex', 'utf8');
            decrypted += legacyDecipher.final('utf8');
            return decrypted;
        } catch (error: any) {
            console.error('[Encryption] Decryption failed:', error.message);
            throw new Error('Failed to decrypt credentials. The encryption key may have changed. Please delete and re-add your email credentials.');
        }
    }
};
