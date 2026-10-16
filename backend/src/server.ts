import express, { Application } from 'express';
import http from 'http';
import { Server as SocketIOServer } from 'socket.io';
import cookieParser from 'cookie-parser';
import { env } from './config/env';
import prisma from './config/database';
import { errorHandler, notFound } from './middleware/errorHandler';
import { apiLimiter } from './middleware/rateLimit';
import { securityHeaders, corsConfig, apiSecurityHeaders } from './middleware/security';
import { requestIdMiddleware, requestLoggerMiddleware } from './middleware/requestLogger';
import { nativeCompressionMiddleware } from './middleware/compression';
import logger from './utils/logger';
import { verifyAccessToken } from './utils/jwt';
import campaignRunManager from './services/campaignRun.manager';
import emailService from './services/email.service';

// Import routes
import authRoutes from './routes/auth.routes';
import credentialRoutes from './routes/credential.routes';
import campaignRoutes from './routes/campaign.routes';
import trackingRoutes from './routes/tracking.routes';
import templateRoutes from './routes/template.routes';
import healthRoutes from './routes/health.routes';

const app: Application = express();
const server = http.createServer(app);

// Socket.IO — attach to the same HTTP server
export const io = new SocketIOServer(server, {
    cors: {
        origin: env.FRONTEND_URL || 'http://localhost:3000',
        methods: ['GET', 'POST'],
        credentials: true,
    },
    transports: ['websocket', 'polling'],
});

// Socket.IO authentication middleware
io.use((socket, next) => {
    try {
        const token = socket.handshake.auth?.token ||
            (socket.handshake.headers.authorization?.startsWith('Bearer ')
                ? socket.handshake.headers.authorization.substring(7)
                : undefined);

        if (!token) {
            return next(new Error('Authentication required'));
        }

        const decoded = verifyAccessToken(token);
        socket.data.user = decoded;
        next();
    } catch (error) {
        next(new Error('Invalid or expired authentication token'));
    }
});

io.on('connection', (socket) => {
    const userId = socket.data.user?.id;
    logger.info(`WebSocket connected: ${socket.id} (user: ${userId})`);

    // Client joins a room for their campaign run
    socket.on('join:campaign', (runId: string) => {
        if (!runId) return;

        const run = campaignRunManager.getRun(runId);
        if (!run || run.userId !== userId) {
            logger.warn(`Unauthorized campaign join attempt: socket ${socket.id} (user ${userId}) tried to join run ${runId}`);
            socket.emit('error', { message: 'Unauthorized: You do not have access to this campaign run' });
            return;
        }

        socket.join(`campaign:${runId}`);
        logger.info(`Socket ${socket.id} joined campaign:${runId}`);
    });

    socket.on('disconnect', () => {
        logger.info(`WebSocket disconnected: ${socket.id}`);
    });
});

// Trust proxy for Render/Railway/Heroku (required for rate limiter + correct IP)
app.set('trust proxy', 1);

// Request ID and logging (must be first)
app.use(requestIdMiddleware);
app.use(requestLoggerMiddleware);

// Security middleware
app.use(securityHeaders);
app.use(corsConfig);
app.use(apiSecurityHeaders);

// Native HTTP response compression (gzip/deflate for payloads > 1KB)
app.use(nativeCompressionMiddleware);

// Cookie parser for CSRF
app.use(cookieParser());

// Body parsing middleware - increased limit for attachments
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Health and Readiness probes (exempt from rate limits)
app.use('/health', healthRoutes);
app.use('/api/health', healthRoutes);

// Rate limiting for application API routes
app.use('/api', apiLimiter);

// API routes
app.use('/api/auth', authRoutes);
app.use('/api/credentials', credentialRoutes);
app.use('/api/campaigns', campaignRoutes);
app.use('/api/track', trackingRoutes);
app.use('/api/templates', templateRoutes);

// Error handling
app.use(notFound);
app.use(errorHandler);

// Start server
const PORT = env.PORT || 5000;

server.listen(PORT, () => {
    logger.info(`Server running on port ${PORT}`);
    logger.info(`Environment: ${env.NODE_ENV}`);
    logger.info(`Frontend URL: ${env.FRONTEND_URL}`);
    logger.info(`Infrastructure: Native Compression + Deep Health Probes + Socket.IO`);

    // Asynchronous database health verification on startup
    prisma.$queryRaw`SELECT 1`
        .then(() => logger.info('Database connection verified on startup.'))
        .catch((err) => logger.error('Database connection verification failed:', err.message));
});

// Graceful Shutdown & Lifecycle Management
const handleShutdown = (signal: string) => {
    logger.info(`Received ${signal}. Initiating graceful shutdown...`);

    // Stop accepting new connections
    server.close(async () => {
        logger.info('HTTP and WebSocket servers closed to new connections.');
        try {
            emailService.clearCache();
            await prisma.$disconnect();
            logger.info('Database and SMTP connections closed cleanly.');
        } catch (err: any) {
            logger.error('Error during cleanup:', err.message);
        }
        logger.info('Graceful shutdown completed successfully.');
        process.exit(0);
    });

    // Forced exit timeout if connections refuse to drain
    setTimeout(() => {
        logger.error('Forced shutdown due to timeout (10s elapsed).');
        process.exit(1);
    }, 10000).unref();
};

process.on('SIGTERM', () => handleShutdown('SIGTERM'));
process.on('SIGINT', () => handleShutdown('SIGINT'));

process.on('unhandledRejection', (reason: any) => {
    logger.error('Unhandled Promise Rejection detected:', reason);
});

process.on('uncaughtException', (error: Error) => {
    logger.error('Uncaught Exception detected:', error);
});

export default app;

