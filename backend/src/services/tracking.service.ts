/**
 * Email Tracking Service
 * Handles open tracking (pixel) and click tracking (link rewriting)
 */
import crypto from 'crypto';
import logger from '../utils/logger';

// In-memory storage for development (file-based)
// In production, this will be handled by database
import fs from 'fs';
import path from 'path';

const TRACKING_DIR = path.join(process.cwd(), 'logs');
const TRACKING_FILE = path.join(TRACKING_DIR, 'email_tracking.jsonl');

// Ensure logs directory exists
if (!fs.existsSync(TRACKING_DIR)) {
    fs.mkdirSync(TRACKING_DIR, { recursive: true });
}

interface TrackingRecord {
    id: string;
    trackingToken: string;
    recipientEmail: string;
    subject?: string;
    campaignId?: string;
    openCount: number;
    firstOpenedAt?: string;
    lastOpenedAt?: string;
    userAgent?: string;
    ipAddress?: string;
    links: LinkRecord[];
    createdAt: string;
}

interface LinkRecord {
    clickToken: string;
    originalUrl: string;
    clickCount: number;
    firstClickedAt?: string;
    lastClickedAt?: string;
}

/**
 * Generate a unique tracking token
 */
export function generateTrackingToken(): string {
    return crypto.randomBytes(16).toString('hex');
}

/**
 * Generate a unique click token for a link
 */
export function generateClickToken(): string {
    return crypto.randomBytes(12).toString('hex');
}

/**
 * Create a tracking record for an email
 */
export async function createTrackingRecord(data: {
    recipientEmail: string;
    subject?: string;
    campaignId?: string;
}): Promise<{ trackingToken: string; trackingPixelUrl: string }> {
    const trackingToken = generateTrackingToken();

    const record: TrackingRecord = {
        id: crypto.randomUUID(),
        trackingToken,
        recipientEmail: data.recipientEmail,
        subject: data.subject,
        campaignId: data.campaignId,
        openCount: 0,
        links: [],
        createdAt: new Date().toISOString(),
    };

    // Append to JSONL file
    const line = JSON.stringify(record) + '\n';
    fs.appendFileSync(TRACKING_FILE, line);

    logger.debug('Created tracking record', { trackingToken, recipientEmail: data.recipientEmail });

    // Generate tracking pixel URL (will be served by our API)
    const baseUrl = process.env.BACKEND_URL || 'http://localhost:5000';
    const trackingPixelUrl = `${baseUrl}/api/track/open/${trackingToken}`;

    return { trackingToken, trackingPixelUrl };
}

/**
 * Record an email open event
 */
export async function recordEmailOpen(
    trackingToken: string,
    userAgent?: string,
    ipAddress?: string
): Promise<boolean> {
    try {
        const records = loadTrackingRecords();
        const recordIndex = records.findIndex(r => r.trackingToken === trackingToken);

        if (recordIndex === -1) {
            logger.warn('Tracking token not found', { trackingToken });
            return false;
        }

        const record = records[recordIndex];
        record.openCount += 1;
        record.lastOpenedAt = new Date().toISOString();
        if (!record.firstOpenedAt) {
            record.firstOpenedAt = record.lastOpenedAt;
        }
        record.userAgent = userAgent;
        record.ipAddress = ipAddress;

        // Save updated records
        saveTrackingRecords(records);

        logger.info('Email opened', {
            trackingToken,
            recipientEmail: record.recipientEmail,
            openCount: record.openCount
        });

        return true;
    } catch (error: any) {
        logger.error('Error recording email open', { error: error.message });
        return false;
    }
}

/**
 * Rewrite links in email body to use tracking URLs
 */
export function rewriteLinksForTracking(
    emailBody: string,
    trackingToken: string,
    baseUrl: string = process.env.BACKEND_URL || 'http://localhost:5000'
): { body: string; links: Array<{ clickToken: string; originalUrl: string }> } {
    const links: Array<{ clickToken: string; originalUrl: string }> = [];

    // Match href attributes in anchor tags
    const linkRegex = /href=["']([^"']+)["']/gi;

    const newBody = emailBody.replace(linkRegex, (match, url) => {
        // Skip mailto:, tel:, and anchor links
        if (url.startsWith('mailto:') || url.startsWith('tel:') || url.startsWith('#')) {
            return match;
        }

        // Skip tracking pixel URLs
        if (url.includes('/api/track/')) {
            return match;
        }

        const clickToken = generateClickToken();
        const trackingUrl = `${baseUrl}/api/track/click/${clickToken}`;

        links.push({ clickToken, originalUrl: url });

        return `href="${trackingUrl}"`;
    });

    // Store link mappings
    if (links.length > 0) {
        const records = loadTrackingRecords();
        const recordIndex = records.findIndex(r => r.trackingToken === trackingToken);
        if (recordIndex !== -1) {
            records[recordIndex].links = links.map(l => ({
                ...l,
                clickCount: 0,
            }));
            saveTrackingRecords(records);
        }
    }

    return { body: newBody, links };
}

/**
 * Record a link click event
 */
export async function recordLinkClick(
    clickToken: string,
    userAgent?: string,
    ipAddress?: string
): Promise<string | null> {
    try {
        const records = loadTrackingRecords();

        for (const record of records) {
            const linkIndex = record.links.findIndex(l => l.clickToken === clickToken);
            if (linkIndex !== -1) {
                const link = record.links[linkIndex];
                link.clickCount += 1;
                link.lastClickedAt = new Date().toISOString();
                if (!link.firstClickedAt) {
                    link.firstClickedAt = link.lastClickedAt;
                }

                saveTrackingRecords(records);

                logger.info('Link clicked', {
                    clickToken,
                    recipientEmail: record.recipientEmail,
                    originalUrl: link.originalUrl,
                    clickCount: link.clickCount,
                });

                return link.originalUrl;
            }
        }

        logger.warn('Click token not found', { clickToken });
        return null;
    } catch (error: any) {
        logger.error('Error recording link click', { error: error.message });
        return null;
    }
}

/**
 * Get tracking stats for a campaign or all emails
 */
export function getTrackingStats(campaignId?: string): {
    totalSent: number;
    totalOpened: number;
    openRate: number;
    totalClicks: number;
    uniqueClicks: number;
    records: TrackingRecord[];
} {
    let records = loadTrackingRecords();

    if (campaignId) {
        records = records.filter(r => r.campaignId === campaignId);
    }

    const totalSent = records.length;
    const totalOpened = records.filter(r => r.openCount > 0).length;
    const openRate = totalSent > 0 ? Math.round((totalOpened / totalSent) * 100) : 0;

    let totalClicks = 0;
    let uniqueClicks = 0;

    for (const record of records) {
        for (const link of record.links) {
            totalClicks += link.clickCount;
            if (link.clickCount > 0) {
                uniqueClicks += 1;
            }
        }
    }

    return {
        totalSent,
        totalOpened,
        openRate,
        totalClicks,
        uniqueClicks,
        records,
    };
}

/**
 * Generate HTML for tracking pixel
 */
export function generateTrackingPixelHtml(trackingToken: string): string {
    const baseUrl = process.env.BACKEND_URL || 'http://localhost:5000';
    const pixelUrl = `${baseUrl}/api/track/open/${trackingToken}`;

    return `<img src="${pixelUrl}" alt="" width="1" height="1" style="display:none;width:1px;height:1px;border:0;" />`;
}

// Helper functions for file-based storage
function loadTrackingRecords(): TrackingRecord[] {
    try {
        if (!fs.existsSync(TRACKING_FILE)) {
            return [];
        }
        const content = fs.readFileSync(TRACKING_FILE, 'utf-8');
        const lines = content.trim().split('\n').filter(Boolean);
        return lines.map(line => JSON.parse(line));
    } catch {
        return [];
    }
}

function saveTrackingRecords(records: TrackingRecord[]): void {
    const content = records.map(r => JSON.stringify(r)).join('\n') + '\n';
    fs.writeFileSync(TRACKING_FILE, content);
}

export default {
    generateTrackingToken,
    generateClickToken,
    createTrackingRecord,
    recordEmailOpen,
    recordLinkClick,
    rewriteLinksForTracking,
    getTrackingStats,
    generateTrackingPixelHtml,
};
