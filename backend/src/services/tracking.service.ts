/**
 * Email Delivery & Outreach Service
 * Records email delivery logs and performance metrics without invasive
 * tracking pixels or link redirects, ensuring optimal deliverability.
 */
import crypto from 'crypto';
import fs from 'fs';
import path from 'path';
import logger from '../utils/logger';

const LOGS_DIR = path.join(process.cwd(), 'logs');
const DELIVERY_LOG_FILE = path.join(LOGS_DIR, 'email_delivery.jsonl');
const DELIVERY_LOG_TMP = path.join(LOGS_DIR, 'email_delivery.tmp');

// Ensure logs directory exists
if (!fs.existsSync(LOGS_DIR)) {
    fs.mkdirSync(LOGS_DIR, { recursive: true });
}

export interface DeliveryRecord {
    id: string;
    userId: string;
    recipientEmail: string;
    subject?: string;
    campaignId?: string;
    status: 'sent' | 'failed';
    errorMessage?: string;
    sentAt: string;
}

// In-memory cache for fast lookups
let recordsCache: DeliveryRecord[] | null = null;
let lastLoadedMtime: number = 0;

/**
 * Load delivery records safely line-by-line.
 */
function loadDeliveryRecords(forceReload = false): DeliveryRecord[] {
    try {
        if (!fs.existsSync(DELIVERY_LOG_FILE)) {
            recordsCache = [];
            lastLoadedMtime = 0;
            return recordsCache;
        }

        const stat = fs.statSync(DELIVERY_LOG_FILE);
        if (!forceReload && recordsCache !== null && stat.mtimeMs === lastLoadedMtime) {
            return recordsCache;
        }

        lastLoadedMtime = stat.mtimeMs;
        const content = fs.readFileSync(DELIVERY_LOG_FILE, 'utf-8');
        const lines = content.split('\n');
        const records: DeliveryRecord[] = [];

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;

            try {
                const parsed = JSON.parse(line);
                if (parsed && parsed.recipientEmail) {
                    // Backwards compatibility for pre-existing logs without userId
                    if (!parsed.userId) {
                        parsed.userId = 'legacy';
                    }
                    records.push(parsed);
                }
            } catch (err: any) {
                logger.warn('Skipping corrupted line in delivery records file', {
                    lineIndex: i,
                    error: err.message,
                });
            }
        }

        recordsCache = records;
        return recordsCache;
    } catch (error: any) {
        logger.error('Failed to read delivery records file', { error: error.message });
        recordsCache = [];
        return recordsCache;
    }
}

/**
 * Save delivery records atomically to prevent file corruption
 */
function saveDeliveryRecords(records: DeliveryRecord[]): void {
    try {
        recordsCache = records;
        const content = records.map(r => JSON.stringify(r)).join('\n') + '\n';

        fs.writeFileSync(DELIVERY_LOG_TMP, content, 'utf-8');
        fs.renameSync(DELIVERY_LOG_TMP, DELIVERY_LOG_FILE);

        if (fs.existsSync(DELIVERY_LOG_FILE)) {
            lastLoadedMtime = fs.statSync(DELIVERY_LOG_FILE).mtimeMs;
        }
    } catch (error: any) {
        logger.error('Failed to save delivery records', { error: error.message });
    }
}

/**
 * Record an email delivery event (sent or failed) scoped to a user
 */
export async function recordDelivery(data: {
    userId: string;
    recipientEmail: string;
    subject?: string;
    campaignId?: string;
    status: 'sent' | 'failed';
    errorMessage?: string;
}): Promise<DeliveryRecord> {
    const record: DeliveryRecord = {
        id: crypto.randomUUID(),
        userId: data.userId,
        recipientEmail: data.recipientEmail,
        subject: data.subject,
        campaignId: data.campaignId,
        status: data.status,
        errorMessage: data.errorMessage,
        sentAt: new Date().toISOString(),
    };

    const records = loadDeliveryRecords();
    records.push(record);
    saveDeliveryRecords(records);

    logger.info(`Email delivery recorded: [${data.status.toUpperCase()}]`, {
        userId: data.userId,
        recipientEmail: data.recipientEmail,
        status: data.status,
    });

    return record;
}

/**
 * Get delivery stats for a specific user (and optionally a specific campaign)
 */
export function getTrackingStats(userId: string, campaignId?: string): {
    totalSent: number;
    sentCount: number;
    failedCount: number;
    successRate: number;
    records: DeliveryRecord[];
} {
    let records = loadDeliveryRecords().filter(r => r.userId === userId);

    if (campaignId) {
        records = records.filter(r => r.campaignId === campaignId);
    }

    const totalSent = records.length;
    const sentCount = records.filter(r => r.status === 'sent').length;
    const failedCount = records.filter(r => r.status === 'failed').length;
    const successRate = totalSent > 0 ? Math.round((sentCount / totalSent) * 100) : 0;

    return {
        totalSent,
        sentCount,
        failedCount,
        successRate,
        records,
    };
}

/**
 * Get paginated delivery details for a specific user (and optionally a specific campaign)
 */
export function getTrackingDetailsPaginated(
    userId: string,
    page: number = 1,
    limit: number = 50,
    campaignId?: string
): {
    totalSent: number;
    sentCount: number;
    failedCount: number;
    successRate: number;
    records: DeliveryRecord[];
    pagination: {
        total: number;
        page: number;
        limit: number;
        totalPages: number;
    };
} {
    let records = loadDeliveryRecords().filter(r => r.userId === userId);

    if (campaignId) {
        records = records.filter(r => r.campaignId === campaignId);
    }

    // Sort by sentAt descending (newest first)
    records.sort((a, b) => new Date(b.sentAt).getTime() - new Date(a.sentAt).getTime());

    const totalSent = records.length;
    const sentCount = records.filter(r => r.status === 'sent').length;
    const failedCount = records.filter(r => r.status === 'failed').length;
    const successRate = totalSent > 0 ? Math.round((sentCount / totalSent) * 100) : 0;

    const startIndex = (page - 1) * limit;
    const endIndex = page * limit;
    const paginatedRecords = records.slice(startIndex, endIndex);

    return {
        totalSent,
        sentCount,
        failedCount,
        successRate,
        records: paginatedRecords,
        pagination: {
            total: totalSent,
            page,
            limit,
            totalPages: Math.ceil(totalSent / limit) || 1,
        },
    };
}

/**
 * Clear delivery records for a specific user (all or for a specific campaign)
 */
export function clearTrackingRecords(userId: string, campaignId?: string): void {
    try {
        let records = loadDeliveryRecords();
        if (campaignId) {
            records = records.filter(r => !(r.userId === userId && r.campaignId === campaignId));
            logger.info('Cleared delivery records for campaign', { userId, campaignId });
        } else {
            records = records.filter(r => r.userId !== userId);
            logger.info('Cleared all email delivery records for user', { userId });
        }
        saveDeliveryRecords(records);
    } catch (error: any) {
        logger.error('Failed to clear delivery records', { error: error.message });
        throw error;
    }
}

// Backward-compatible stubs so no external calls break
export function generateTrackingToken(): string {
    return '';
}
export function generateClickToken(): string {
    return '';
}
export async function createTrackingRecord(): Promise<{ trackingToken: string; trackingPixelUrl: string }> {
    return { trackingToken: '', trackingPixelUrl: '' };
}
export function deleteTrackingRecord(_trackingToken?: string): void {}
export async function recordEmailOpen(_token?: string, _userAgent?: string, _ipAddress?: string): Promise<boolean> {
    return false;
}
export async function recordLinkClick(_token?: string, _userAgent?: string, _ipAddress?: string): Promise<string | null> {
    return null;
}
export function rewriteLinksForTracking(emailBody: string): { body: string; links: any[] } {
    return { body: emailBody, links: [] };
}
export function generateTrackingPixelHtml(): string {
    return '';
}

export default {
    recordDelivery,
    getTrackingStats,
    getTrackingDetailsPaginated,
    clearTrackingRecords,
    generateTrackingToken,
    generateClickToken,
    createTrackingRecord,
    deleteTrackingRecord,
    recordEmailOpen,
    recordLinkClick,
    rewriteLinksForTracking,
    generateTrackingPixelHtml,
};
