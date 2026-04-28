import nodemailer, { Transporter } from 'nodemailer';
import { AppError } from '../middleware/errorHandler';
import logger from '../utils/logger';
import trackingService from './tracking.service';

interface EmailOptions {
    from: string;
    to: string;
    subject: string;
    html: string;
    attachments?: Array<{
        filename: string;
        path?: string;
        content?: Buffer;
        contentType?: string;
    }>;
}

interface SMTPConfig {
    email: string;
    appPassword: string;
}

interface CachedTransporter {
    transporter: Transporter;
    lastUsed: number;
    verified: boolean;
}

/**
 * EmailService with connection pooling for improved performance.
 * Transporters are cached per email account and reused across requests.
 */
export class EmailService {
    private transporterCache: Map<string, CachedTransporter> = new Map();
    private readonly CACHE_TTL = 30 * 60 * 1000; // 30 minutes
    private readonly MAX_CONNECTIONS = 5;

    /**
     * Gets or creates a pooled transporter for the given email account
     */
    private async getTransporter(config: SMTPConfig): Promise<Transporter> {
        const cacheKey = config.email;
        const cached = this.transporterCache.get(cacheKey);
        const now = Date.now();

        // Return cached transporter if valid
        if (cached && cached.verified && (now - cached.lastUsed) < this.CACHE_TTL) {
            cached.lastUsed = now;
            logger.debug(`Using cached transporter for ${config.email}`);
            return cached.transporter;
        }

        // Close old transporter if exists
        if (cached) {
            try {
                cached.transporter.close();
            } catch {
                // Ignore close errors
            }
            this.transporterCache.delete(cacheKey);
        }

        logger.info(`Creating new pooled transporter for ${config.email}`);

        // Create pooled transporter
        const transporter = nodemailer.createTransport({
            host: 'smtp.gmail.com',
            port: 587,
            secure: false,
            pool: true, // Enable connection pooling
            maxConnections: this.MAX_CONNECTIONS,
            maxMessages: 100, // Messages per connection before reconnect
            auth: {
                user: config.email,
                pass: config.appPassword,
            },
            // Disable verbose logging in production
            debug: process.env.NODE_ENV === 'development',
            logger: process.env.NODE_ENV === 'development',
        });

        // Verify connection
        try {
            await transporter.verify();
            logger.info(`SMTP connection verified for ${config.email}`);
        } catch (verifyError: any) {
            logger.error(`SMTP verification failed for ${config.email}: ${verifyError.message}`);
            throw new AppError(`SMTP authentication failed: ${verifyError.message}`, 401);
        }

        // Cache the transporter
        this.transporterCache.set(cacheKey, {
            transporter,
            lastUsed: now,
            verified: true,
        });

        return transporter;
    }

    /**
     * Sends an email using a pooled connection
     */
    async sendEmail(config: SMTPConfig, options: EmailOptions): Promise<void> {
        logger.info(`Sending email to: ${options.to} from: ${config.email}`);

        try {
            const transporter = await this.getTransporter(config);

            if (options.attachments?.length) {
                logger.debug(`Email has ${options.attachments.length} attachment(s)`);
            }

            const info = await transporter.sendMail({
                from: options.from,
                to: options.to,
                subject: options.subject,
                html: options.html,
                attachments: options.attachments,
            });

            logger.info(`Email sent successfully! MessageId: ${info.messageId}`);
        } catch (error: any) {
            logger.error(`Email sending failed: ${error.message}`, {
                to: options.to,
                errorCode: error.code,
            });

            // Invalidate cached transporter on auth errors
            if (error.code === 'EAUTH') {
                this.transporterCache.delete(config.email);
            }

            // Provide user-friendly error messages
            if (error.code === 'EAUTH' || error.message?.includes('Invalid login')) {
                throw new AppError('Gmail authentication failed. Please check your App Password.', 401);
            } else if (error.code === 'ESOCKET' || error.code === 'ECONNECTION') {
                throw new AppError('Unable to connect to Gmail. Please check your internet connection.', 503);
            } else if (error instanceof AppError) {
                throw error;
            } else {
                throw new AppError(`Email sending failed: ${error.message}`, 500);
            }
        }
    }

    /**
     * Sends multiple emails efficiently using the pooled connection
     */
    async sendBulkEmails(
        config: SMTPConfig,
        emails: EmailOptions[],
        onProgress?: (sent: number, total: number) => void
    ): Promise<{ sent: number; failed: number; errors: string[] }> {
        const results = { sent: 0, failed: 0, errors: [] as string[] };

        for (let i = 0; i < emails.length; i++) {
            try {
                await this.sendEmail(config, emails[i]);
                results.sent++;
            } catch (error: any) {
                results.failed++;
                results.errors.push(`${emails[i].to}: ${error.message}`);
            }

            if (onProgress) {
                onProgress(i + 1, emails.length);
            }
        }

        return results;
    }

    /**
     * Personalizes email content with recipient data
     */
    personalizeContent(template: string, data: Record<string, string>): string {
        let result = template;
        for (const [key, value] of Object.entries(data)) {
            const placeholder = new RegExp(`\\{${key}\\}`, 'gi');
            result = result.replace(placeholder, value || '');
        }
        return result;
    }

    /**
     * Prepares an email with tracking (pixel + link rewriting)
     * Call this before sending to enable open and click tracking
     */
    async prepareTrackedEmail(options: {
        recipientEmail: string;
        subject: string;
        html: string;
        campaignId?: string;
    }): Promise<{ html: string; trackingToken: string }> {
        try {
            // Create tracking record
            const { trackingToken } = await trackingService.createTrackingRecord({
                recipientEmail: options.recipientEmail,
                subject: options.subject,
                campaignId: options.campaignId,
            });

            // Rewrite links for click tracking
            const { body: bodyWithTrackedLinks } = trackingService.rewriteLinksForTracking(
                options.html,
                trackingToken
            );

            // Generate and append tracking pixel
            const trackingPixel = trackingService.generateTrackingPixelHtml(trackingToken);

            // Inject pixel before closing body tag, or at the end
            let trackedHtml = bodyWithTrackedLinks;
            if (trackedHtml.includes('</body>')) {
                trackedHtml = trackedHtml.replace('</body>', `${trackingPixel}</body>`);
            } else {
                trackedHtml = trackedHtml + trackingPixel;
            }

            logger.debug('Email prepared with tracking', {
                recipientEmail: options.recipientEmail,
                trackingToken
            });

            return { html: trackedHtml, trackingToken };
        } catch (error: any) {
            // If tracking fails, still return original HTML
            logger.error('Failed to prepare tracked email', { error: error.message });
            return { html: options.html, trackingToken: '' };
        }
    }

    /**
     * Clears all cached transporters (useful for cleanup)
     */
    clearCache(): void {
        for (const [email, cached] of this.transporterCache) {
            try {
                cached.transporter.close();
                logger.debug(`Closed transporter for ${email}`);
            } catch {
                // Ignore close errors
            }
        }
        this.transporterCache.clear();
        logger.info('Email transporter cache cleared');
    }

    /**
     * Gets cache statistics
     */
    getCacheStats(): { size: number; emails: string[] } {
        return {
            size: this.transporterCache.size,
            emails: Array.from(this.transporterCache.keys()),
        };
    }
}

export default new EmailService();
