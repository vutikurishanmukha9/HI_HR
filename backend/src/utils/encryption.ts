import crypto from 'crypto';
import { env } from '../config/env';

const ALGORITHM = 'aes-256-cbc';
const IV_LENGTH = 16;

/**
 * Encrypts sensitive data using AES-256-CBC
 */
export const encrypt = (text: string): string => {
    const iv = crypto.randomBytes(IV_LENGTH);
    const cipher = crypto.createCipheriv(
        ALGORITHM,
        Buffer.from(env.ENCRYPTION_KEY, 'utf-8').slice(0, 32),
        iv
    );

    let encrypted = cipher.update(text, 'utf8', 'hex');
    encrypted += cipher.final('hex');

    return iv.toString('hex') + ':' + encrypted;
};

/**
 * Decrypts data encrypted with the encrypt function
 */
export const decrypt = (text: string): string => {
    try {
        const parts = text.split(':');
        if (parts.length < 2) {
            throw new Error('Invalid encrypted data format');
        }

        const iv = Buffer.from(parts.shift()!, 'hex');
        const encryptedText = parts.join(':');

        const decipher = crypto.createDecipheriv(
            ALGORITHM,
            Buffer.from(env.ENCRYPTION_KEY, 'utf-8').slice(0, 32),
            iv
        );

        let decrypted = decipher.update(encryptedText, 'hex', 'utf8');
        decrypted += decipher.final('utf8');

        return decrypted;
    } catch (error: any) {
        console.error('[Encryption] Decryption failed:', error.message);
        throw new Error('Failed to decrypt credentials. The encryption key may have changed. Please delete and re-add your email credentials.');
    }
};
