/**
 * Email Tracking Routes
 * Handles tracking pixel requests and link click redirects
 */
import { Router, Response } from 'express';
import { recordEmailOpen, recordLinkClick, getTrackingStats, getTrackingDetailsPaginated, clearTrackingRecords } from '../services/tracking.service';
import { authenticate } from '../middleware/auth';
import { AuthRequest } from '../types';
import campaignRunManager from '../services/campaignRun.manager';
import prisma from '../config/database';
import logger from '../utils/logger';

const router = Router();

// Helper to verify campaign ownership for a user
async function verifyCampaignOwnership(userId: string, campaignId: string): Promise<boolean> {
    const run = campaignRunManager.getRun(campaignId);
    if (run && run.userId === userId) {
        return true;
    }
    const dbCampaign = await prisma.campaign.findFirst({
        where: { id: campaignId, userId },
        select: { id: true },
    });
    return !!dbCampaign;
}

// 1x1 transparent GIF pixel
const TRANSPARENT_PIXEL = Buffer.from(
    'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
    'base64'
);

/**
 * GET /api/track/open/:token
 * Tracking pixel endpoint - records email opens
 * Returns a 1x1 transparent GIF with anti-caching and CORS headers
 */
router.all('/open/:token', async (req: AuthRequest, res: Response) => {
    const token = req.params.token?.trim();
    const userAgent = req.headers['user-agent'];
    const ipAddress = req.ip || req.socket.remoteAddress;

    // Set bulletproof headers for email clients and image proxies (e.g. Google Image Proxy)
    res.set({
        'Content-Type': 'image/gif',
        'Content-Length': TRANSPARENT_PIXEL.length.toString(),
        'Cache-Control': 'no-store, no-cache, must-revalidate, proxy-revalidate, max-age=0, post-check=0, pre-check=0',
        'Pragma': 'no-cache',
        'Expires': '0',
        'Surrogate-Control': 'no-store',
        'Access-Control-Allow-Origin': '*',
        'Cross-Origin-Resource-Policy': 'cross-origin',
        'Timing-Allow-Origin': '*',
    });

    try {
        // Only count actual GET requests (ignore HEAD requests from prefetchers/scanners)
        if (req.method === 'GET' && token) {
            recordEmailOpen(token, userAgent, ipAddress).catch(err => {
                logger.error('Failed to record email open', { error: err.message });
            });
        }

        res.status(200).end(TRANSPARENT_PIXEL);
    } catch (error) {
        // Always return the pixel even if error occurs
        res.status(200).end(TRANSPARENT_PIXEL);
    }
});

/**
 * GET /api/track/click/:token
 * Link click tracking - records click and redirects to original URL
 */
router.get('/click/:token', async (req: AuthRequest, res: Response) => {
    const token = req.params.token?.trim();
    const userAgent = req.headers['user-agent'];
    const ipAddress = req.ip || req.socket.remoteAddress;

    if (!token) {
        res.status(400).send('Invalid click token');
        return;
    }

    try {
        const originalUrl = await recordLinkClick(token, userAgent, ipAddress);

        if (originalUrl) {
            // Redirect with 302 to destination URL
            res.redirect(302, originalUrl);
        } else {
            // Token not found or expired
            logger.warn('Click tracking link not found', { token });
            res.status(404).send('Link not found or expired');
        }
    } catch (error: any) {
        logger.error('Error processing click tracking', { error: error.message });
        res.status(500).send('Error processing link redirect');
    }
});

/**
 * GET /api/track/stats
 * Get delivery statistics (authenticated and scoped to current user)
 */
router.get('/stats', authenticate, async (req: AuthRequest, res: Response) => {
    try {
        const userId = req.user!.id;
        const { campaignId } = req.query;
        const targetCampaignId = campaignId ? String(campaignId) : undefined;

        if (targetCampaignId) {
            const isOwner = await verifyCampaignOwnership(userId, targetCampaignId);
            if (!isOwner) {
                res.status(403).json({ success: false, error: 'Access denied: You do not own this campaign' });
                return;
            }
        }

        const stats = getTrackingStats(userId, targetCampaignId);

        res.json({
            success: true,
            data: {
                totalSent: stats.totalSent,
                sentCount: stats.sentCount,
                failedCount: stats.failedCount,
                successRate: stats.successRate,
            },
        });
    } catch (error: any) {
        logger.error('Error getting delivery stats', { error: error.message });
        res.status(500).json({ success: false, error: 'Failed to get delivery stats' });
    }
});

/**
 * GET /api/track/details
 * Get detailed delivery records (authenticated and scoped to current user)
 */
router.get('/details', authenticate, async (req: AuthRequest, res: Response) => {
    try {
        const userId = req.user!.id;
        const { campaignId, page = '1', limit = '50' } = req.query;
        const targetCampaignId = campaignId ? String(campaignId) : undefined;

        if (targetCampaignId) {
            const isOwner = await verifyCampaignOwnership(userId, targetCampaignId);
            if (!isOwner) {
                res.status(403).json({ success: false, error: 'Access denied: You do not own this campaign' });
                return;
            }
        }

        const pageNum = parseInt(page as string, 10) || 1;
        const limitNum = parseInt(limit as string, 10) || 50;

        const stats = getTrackingDetailsPaginated(userId, pageNum, limitNum, targetCampaignId);

        // Transform records for frontend
        const details = stats.records.map(record => ({
            recipientEmail: record.recipientEmail,
            subject: record.subject,
            campaignId: record.campaignId,
            status: record.status,
            sentAt: record.sentAt,
            errorMessage: record.errorMessage,
        }));

        res.json({
            success: true,
            data: {
                summary: {
                    totalSent: stats.totalSent,
                    sentCount: stats.sentCount,
                    failedCount: stats.failedCount,
                    successRate: stats.successRate,
                },
                details,
                pagination: stats.pagination
            },
        });
    } catch (error: any) {
        logger.error('Error getting delivery details', { error: error.message });
        res.status(500).json({ success: false, error: 'Failed to get delivery details' });
    }
});

/**
 * POST /api/track/reset
 * Reset/clear tracking statistics (authenticated and scoped to current user)
 */
router.post('/reset', authenticate, async (req: AuthRequest, res: Response) => {
    try {
        const userId = req.user!.id;
        const { campaignId } = req.query;
        const targetCampaignId = campaignId ? String(campaignId) : undefined;

        if (targetCampaignId) {
            const isOwner = await verifyCampaignOwnership(userId, targetCampaignId);
            if (!isOwner) {
                res.status(403).json({ success: false, error: 'Access denied: You do not own this campaign' });
                return;
            }
        }

        clearTrackingRecords(userId, targetCampaignId);
        logger.info('Tracking stats reset via POST', { userId, campaignId: targetCampaignId });
        res.json({
            success: true,
            message: 'Tracking statistics cleared successfully',
        });
    } catch (error: any) {
        logger.error('Error resetting tracking stats', { error: error.message });
        res.status(500).json({ success: false, error: 'Failed to reset tracking stats' });
    }
});

/**
 * DELETE /api/track/reset
 * Reset/clear tracking statistics (authenticated and scoped to current user)
 */
router.delete('/reset', authenticate, async (req: AuthRequest, res: Response) => {
    try {
        const userId = req.user!.id;
        const { campaignId } = req.query;
        const targetCampaignId = campaignId ? String(campaignId) : undefined;

        if (targetCampaignId) {
            const isOwner = await verifyCampaignOwnership(userId, targetCampaignId);
            if (!isOwner) {
                res.status(403).json({ success: false, error: 'Access denied: You do not own this campaign' });
                return;
            }
        }

        clearTrackingRecords(userId, targetCampaignId);
        logger.info('Tracking stats reset via DELETE', { userId, campaignId: targetCampaignId });
        res.json({
            success: true,
            message: 'Tracking statistics cleared successfully',
        });
    } catch (error: any) {
        logger.error('Error resetting tracking stats', { error: error.message });
        res.status(500).json({ success: false, error: 'Failed to reset tracking stats' });
    }
});

export default router;
