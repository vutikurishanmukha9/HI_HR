/**
 * Unit tests for encryption utilities
 * Tests AES-256-CBC encryption and decryption
 */

// Mock the env module
jest.mock('../../../src/config/env', () => ({
    env: {
        ENCRYPTION_KEY: '12345678901234567890123456789012', // Exactly 32 chars
        NODE_ENV: 'test',
    }
}));

import { encrypt, decrypt } from '../../../src/utils/encryption';

describe('Encryption Utilities', () => {
    describe('encrypt', () => {
        it('should encrypt a string', () => {
            const plainText = 'my-secret-password';
            const encrypted = encrypt(plainText);

            expect(encrypted).toBeDefined();
            expect(typeof encrypted).toBe('string');
            expect(encrypted).not.toBe(plainText);
        });

        it('should produce different output for same input (due to random IV)', () => {
            const plainText = 'my-secret-password';
            const encrypted1 = encrypt(plainText);
            const encrypted2 = encrypt(plainText);

            expect(encrypted1).not.toBe(encrypted2);
        });

        it('should include IV separator in output', () => {
            const encrypted = encrypt('test');
            expect(encrypted).toContain(':');
        });

        it('should handle empty strings', () => {
            const encrypted = encrypt('');
            expect(encrypted).toBeDefined();
            expect(encrypted).toContain(':');
        });

        it('should handle special characters', () => {
            const specialText = 'p@$$w0rd!#$%^&*()_+-={}[]|:;<>?,./~`';
            const encrypted = encrypt(specialText);
            expect(encrypted).toBeDefined();
        });

        it('should handle unicode characters', () => {
            const unicodeText = '密码 пароль 🔐';
            const encrypted = encrypt(unicodeText);
            expect(encrypted).toBeDefined();
        });
    });

    describe('decrypt', () => {
        it('should decrypt an encrypted string back to original', () => {
            const plainText = 'my-secret-password';
            const encrypted = encrypt(plainText);
            const decrypted = decrypt(encrypted);

            expect(decrypted).toBe(plainText);
        });

        it('should handle special characters', () => {
            const specialText = 'p@$$w0rd!#$%^&*()_+-={}[]|:;<>?,./~`';
            const encrypted = encrypt(specialText);
            const decrypted = decrypt(encrypted);

            expect(decrypted).toBe(specialText);
        });

        it('should handle unicode characters', () => {
            const unicodeText = '密码 пароль 🔐';
            const encrypted = encrypt(unicodeText);
            const decrypted = decrypt(encrypted);

            expect(decrypted).toBe(unicodeText);
        });

        it('should handle long strings', () => {
            const longText = 'a'.repeat(10000);
            const encrypted = encrypt(longText);
            const decrypted = decrypt(encrypted);

            expect(decrypted).toBe(longText);
        });

        it('should throw error for invalid format (no separator)', () => {
            expect(() => decrypt('invalid-no-separator')).toThrow();
        });

        it('should throw error for tampered ciphertext', () => {
            const encrypted = encrypt('test');
            const tampered = encrypted.replace(/[0-9a-f]/, 'x');

            expect(() => decrypt(tampered)).toThrow();
        });
    });

    describe('encrypt/decrypt roundtrip', () => {
        const testCases = [
            'simple text',
            '',
            ' ',
            'with\nnewlines\n',
            'with\ttabs',
            JSON.stringify({ email: 'test@example.com', password: 'secret' }),
            'a'.repeat(1000),
        ];

        testCases.forEach((testCase) => {
            it(`should roundtrip: "${testCase.substring(0, 20)}..."`, () => {
                const encrypted = encrypt(testCase);
                const decrypted = decrypt(encrypted);
                expect(decrypted).toBe(testCase);
            });
        });
    });
});
