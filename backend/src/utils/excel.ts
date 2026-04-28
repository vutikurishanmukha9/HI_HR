import XLSX from 'xlsx';
import { RecipientDTO } from '../types';

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
        const fullName = findValue(row, ['Full name', 'Name', 'Last name']);
        const email = findValue(row, ['Email address', 'Email']);
        const companyName = findValue(row, ['Company/organization name', 'Company', 'Organization']);
        const jobTitle = findValue(row, ['Job title/designation', 'Job Title', 'Title', 'Designation']);

        if (fullName && email && companyName) {
            recipients.push({
                fullName,
                email,
                companyName,
                jobTitle: jobTitle || undefined,
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
