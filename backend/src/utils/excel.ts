import XLSX from 'xlsx';
import { RecipientDTO } from '../types';

/**
 * Helper to sanitize header strings to camelCase (e.g. "First Name" -> "firstName")
 */
const sanitizeKey = (key: string): string => {
    const cleaned = key.replace(/[^a-zA-Z0-9]/g, ' ').trim();
    return cleaned.split(/\s+/).map((word, index) => {
        if (index === 0) return word.toLowerCase();
        return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
    }).join('');
};

/**
 * Parses an Excel file and extracts recipient data
 */
export const parseExcelFile = (buffer: Buffer): RecipientDTO[] => {
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);

    const recipients: RecipientDTO[] = [];

    for (const row of data as any[]) {
        // Find the email using common patterns before sanitization
        const email = findValue(row, ['Email address', 'Email', 'E-mail', 'Contact Email']);
        
        if (email) {
            const sanitizedRow: Record<string, any> = {};
            Object.keys(row).forEach(key => {
                const camelKey = sanitizeKey(key);
                if (camelKey) {
                    sanitizedRow[camelKey] = String(row[key]).trim();
                }
            });

            recipients.push({
                ...sanitizedRow,
                email: email // explicitly set the email
            });
        }
    }

    return recipients;
};

/**
 * Helper function to find a value in a row by multiple possible keys
 */
const findValue = (row: any, potentialKeys: string[]): string => {
    const rowKeys = Object.keys(row);
    for (const pKey of potentialKeys) {
        const foundKey = rowKeys.find(
            (rKey) => rKey.trim().toLowerCase() === pKey.trim().toLowerCase()
        );
        if (foundKey && row[foundKey] != null) {
            return String(row[foundKey]).trim();
        }
    }
    return '';
};

/**
 * Removes duplicate recipients based on email address
 */
export const removeDuplicateRecipients = (recipients: RecipientDTO[]): RecipientDTO[] => {
    const seen = new Set<string>();
    return recipients.filter((recipient) => {
        const email = recipient.email.toLowerCase();
        if (seen.has(email)) {
            return false;
        }
        seen.add(email);
        return true;
    });
};
