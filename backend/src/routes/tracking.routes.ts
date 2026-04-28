/**
 * Email Tracking Routes
 * Handles tracking pixel requests and link click redirects
 */
import { Router, Request, Response } from 'express';
import { recordEmailOpen, recordLinkClick, getTrackingStats } from '../services/tracking.service';
import { authenticate } from '../middleware/auth';
import logger from '../utils/logger';

const router = Router();

// 1x1 transparent GIF pixel
const TRANSPARENT_PIXEL = Buffer.from(
    'R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
    'base64'
);

/**
 * GET /api/track/open/:token
 * Tracking pixel endpoint - records email opens
 * Returns a 1x1 transparent GIF
 */
router.get('/open/:token', async (req: Request, res: Response) => {
    const { token } = req.params;
    const userAgent = req.headers['user-agent'];
    const ipAddress = req.ip || req.socket.remoteAddress;

    try {
        // Record the open event (async, don't wait)
        recordEmailOpen(token, userAgent, ipAddress).catch(err => {
            logger.error('Failed to record email open', { error: err.message });
        });

        // Return transparent 1x1 pixel
        res.set({
            'Content-Type': 'image/gif',
            'Content-Length': TRANSPARENT_PIXEL.length,
            'Cache-Control': 'no-store, no-cache, must-revalidate, proxy-revalidate',
            'Pragma': 'no-cache',
            'Expires': '0',
        });
        res.send(TRANSPARENT_PIXEL);
    } catch (error) {
        // Still return the pixel even on error
        res.set('Content-Type', 'image/gif');
        res.send(TRANSPARENT_PIXEL);
    }
});

/**
 * GET /api/track/click/:token
 * Link click tracking - records click and redirects to original URL
 */
router.get('/click/:token', async (req: Request, res: Response) => {
    const { token } = req.params;
    const userAgent = req.headers['user-agent'];
    const ipAddress = req.ip || req.socket.remoteAddress;

    try {
        const originalUrl = await recordLinkClick(token, userAgent, ipAddress);

        if (originalUrl) {
            // Redirect to original URL
            res.redirect(302, originalUrl);
        } else {
            // Token not found - redirect to homepage or show error
            res.status(404).send('Link not found or expired');
        }
    } catch (error: any) {
        logger.error('Error processing click tracking', { error: error.message });
        res.status(500).send('Error processing request');
    }
});

/**
 * GET /api/track/stats
 * Get tracking statistics (authenticated)
 */
router.get('/stats', authenticate, async (req: Request, res: Response) => {
    try {
        const { campaignId } = req.query;
        const stats = getTrackingStats(campaignId as string | undefined);

        res.json({
            success: true,
            data: {
                totalSent: stats.totalSent,
                totalOpened: stats.totalOpened,
                openRate: stats.openRate,
                totalClicks: stats.totalClicks,
                uniqueClicks: stats.uniqueClicks,
            },
        });
    } catch (error: any) {
        logger.error('Error getting tracking stats', { error: error.message });
        res.status(500).json({ success: false, error: 'Failed to get tracking stats' });
    }
});

/**
 * GET /api/track/details
 * Get detailed tracking records (authenticated)
 */
router.get('/details', authenticate, async (req: Request, res: Response) => {
    try {
        const { campaignId } = req.query;
        const stats = getTrackingStats(campaignId as string | undefined);

        // Transform records for frontend
        const details = stats.records.map(record => ({
            recipientEmail: record.recipientEmail,
            subject: record.subject,
            opened: record.openCount > 0,
            openCount: record.openCount,
            firstOpenedAt: record.firstOpenedAt,
            lastOpenedAt: record.lastOpenedAt,
            clicks: record.links.filter(l => l.clickCount > 0).length,
            links: record.links.map(link => ({
                url: link.originalUrl,
                clicks: link.clickCount,
                firstClickedAt: link.firstClickedAt,
            })),
        }));

        res.json({
            success: true,
            data: {
                summary: {
                    totalSent: stats.totalSent,
                    totalOpened: stats.totalOpened,
                    openRate: stats.openRate,
                    totalClicks: stats.totalClicks,
                },
                details,
            },
        });
    } catch (error: any) {
        logger.error('Error getting tracking details', { error: error.message });
        res.status(500).json({ success: false, error: 'Failed to get tracking details' });
    }
});

export default router;
