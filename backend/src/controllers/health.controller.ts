/**
 * Health & Readiness Controller
 * Production-grade telemetry, liveness and readiness probes
 * for Kubernetes, Docker, and Cloud Load Balancer orchestrators.
 */
import { Request, Response } from 'express';
import prisma from '../config/database';
import emailService from '../services/email.service';
import campaignRunManager from '../services/campaignRun.manager';
import { io } from '../server';

export class HealthController {
    /**
     * Liveness Probe — /health or /health/live
     * Fast check confirming process is running and accepting event loop tasks
     */
    static getLiveness(req: Request, res: Response): void {
        res.status(200).json({
            status: 'ok',
            service: 'hihr-backend',
            uptimeSeconds: Math.floor(process.uptime()),
            timestamp: new Date().toISOString(),
            pid: process.pid,
            nodeVersion: process.version,
        });
    }

    /**
     * Readiness Probe — /health/ready
     * Verifies critical dependencies (database latency, socket engine)
     * before routing traffic to this instance.
     */
    static async getReadiness(req: Request, res: Response): Promise<void> {
        const start = Date.now();
        let dbStatus = 'healthy';
        let dbLatencyMs = 0;
        let isReady = true;

        try {
            // Execute quick lightweight ping query
            await prisma.$queryRaw`SELECT 1`;
            dbLatencyMs = Date.now() - start;
        } catch (error: any) {
            dbStatus = 'unreachable';
            dbLatencyMs = Date.now() - start;
            isReady = false;
        }

        const memory = process.memoryUsage();
        const memoryMb = {
            heapUsed: Math.round(memory.heapUsed / 1024 / 1024),
            heapTotal: Math.round(memory.heapTotal / 1024 / 1024),
            rss: Math.round(memory.rss / 1024 / 1024),
        };

        const activeSockets = io?.sockets?.sockets?.size || 0;
        const cacheStats = emailService.getCacheStats();

        const payload = {
            status: isReady ? 'ready' : 'not_ready',
            checks: {
                database: {
                    status: dbStatus,
                    latencyMs: dbLatencyMs,
                },
                smtpPool: {
                    activeCachedTransporters: cacheStats.size,
                },
                websocket: {
                    connectedClients: activeSockets,
                },
            },
            resources: {
                memoryMb,
                uptimeSeconds: Math.floor(process.uptime()),
            },
            timestamp: new Date().toISOString(),
        };

        if (!isReady) {
            res.status(503).json(payload);
            return;
        }

        res.status(200).json(payload);
    }

    /**
     * System Metrics & Diagnostics — /health/metrics
     */
    static async getMetrics(req: Request, res: Response): Promise<void> {
        try {
            const [userCount, campaignCount, recipientCount] = await Promise.all([
                prisma.user.count().catch(() => 0),
                prisma.campaign.count().catch(() => 0),
                prisma.recipient.count().catch(() => 0),
            ]);

            const memory = process.memoryUsage();
            const cacheStats = emailService.getCacheStats();
            const activeSockets = io?.sockets?.sockets?.size || 0;

            res.status(200).json({
                system: {
                    platform: process.platform,
                    arch: process.arch,
                    nodeVersion: process.version,
                    uptimeSeconds: Math.floor(process.uptime()),
                    cpuUsage: process.cpuUsage(),
                },
                memory: {
                    heapUsedMb: Math.round(memory.heapUsed / 1024 / 1024),
                    heapTotalMb: Math.round(memory.heapTotal / 1024 / 1024),
                    rssMb: Math.round(memory.rss / 1024 / 1024),
                    externalMb: Math.round(memory.external / 1024 / 1024),
                },
                connections: {
                    activeWebSockets: activeSockets,
                    cachedSmtpTransporters: cacheStats.size,
                },
                aggregateCounts: {
                    users: userCount,
                    campaigns: campaignCount,
                    recipients: recipientCount,
                },
                timestamp: new Date().toISOString(),
            });
        } catch (error: any) {
            res.status(500).json({
                error: 'Failed to collect diagnostics',
                message: error.message,
            });
        }
    }
}

export default HealthController;
