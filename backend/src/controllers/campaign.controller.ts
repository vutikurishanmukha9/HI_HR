import { Response, NextFunction } from 'express';
import { AuthRequest, RecipientDTO } from '../types';
import campaignService from '../services/campaign.service';
import credentialService from '../services/credential.service';
import emailService from '../services/email.service';
import { parseExcelFile } from '../utils/excel';
import { AppError, ApiError, ErrorCode } from '../middleware/errorHandler';
import multer from 'multer';
import path from 'path';
import { env } from '../config/env';
import campaignRunManager from '../services/campaignRun.manager';
import trackingService from '../services/tracking.service';
import emailValidatorService from '../services/emailValidator.service';
import { io } from '../server';

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, env.UPLOAD_DIR);
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1e9);
        cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
    },
});

export const upload = multer({
    storage,
    limits: { fileSize: env.MAX_FILE_SIZE },
    fileFilter: (req, file, cb) => {
        const allowedTypes = ['.xlsx', '.xls'];
        const ext = path.extname(file.originalname).toLowerCase();
        if (allowedTypes.includes(ext)) {
            cb(null, true);
        } else {
            cb(new Error('Only Excel files are allowed'));
        }
    },
});

interface SendCampaignRequest {
    credentialId?: string;
    credentialEmail?: string;
    subject: string;
    body: string;
    recipients: Array<{
        email: string;
        [key: string]: any;
    }>;
    attachments?: Array<{
        filename: string;
        content: string; // base64 encoded
        contentType: string;
    }>;
    batchSize?: number;
    batchDelay?: number;
    useJitter?: boolean;
}

export class CampaignController {
    async createCampaign(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const recipients: RecipientDTO[] = req.body.recipients || [];
            const campaign = await campaignService.createCampaign(req.user!.id, req.body, recipients);
            res.status(201).json({ campaign });
        } catch (error) {
            next(error);
        }
    }

    async getCampaigns(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const result = await campaignService.getCampaigns(req.user!.id, req.query);
            res.json(result);
        } catch (error) {
            next(error);
        }
    }

    async getCampaignById(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const campaign = await campaignService.getCampaignById(req.user!.id, req.params.id);
            res.json({ campaign });
        } catch (error) {
            next(error);
        }
    }

    async updateCampaign(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const campaign = await campaignService.updateCampaign(req.user!.id, req.params.id, req.body);
            res.json({ campaign });
        } catch (error) {
            next(error);
        }
    }

    async deleteCampaign(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const result = await campaignService.deleteCampaign(req.user!.id, req.params.id);
            res.json(result);
        } catch (error) {
            next(error);
        }
    }

    async uploadRecipients(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            if (!req.file) {
                res.status(400).json({ error: 'No file uploaded' });
                return;
            }

            const fileBuffer = require('fs').readFileSync(req.file.path);
            const recipients = parseExcelFile(fileBuffer);

            // Clean up uploaded file
            require('fs').unlinkSync(req.file.path);

            res.json({ recipients, count: recipients.length });
        } catch (error) {
            // Clean up file on error
            if (req.file) {
                require('fs').unlinkSync(req.file.path);
            }
            next(error);
        }
    }

    /**
     * Send campaign emails using backend Nodemailer
     * Fire-and-forget: returns immediately with a runId,
     * then sends emails in background emitting WebSocket events.
     */
    async sendCampaign(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const {
                credentialEmail,
                subject,
                body,
                recipients,
                attachments,
                batchSize = 10,
                batchDelay = 60,
                useJitter = true
            } = req.body as SendCampaignRequest;

            const userId = req.user!.id;

            if (!recipients || recipients.length === 0) {
                throw ApiError.badRequest('No recipients provided. Please upload at least one recipient.', ErrorCode.MISSING_REQUIRED_FIELD);
            }

            // Strip CR/LF to prevent SMTP header injection
            const cleanSubject = (subject || '').replace(/[\r\n]+/g, ' ').trim();

            if (!cleanSubject || !body) {
                throw ApiError.badRequest('Subject and body are required. Please compose your email before sending.', ErrorCode.MISSING_REQUIRED_FIELD);
            }

            // Validate attachments against dangerous executable extensions
            if (attachments && Array.isArray(attachments)) {
                const dangerousExtensions = ['.exe', '.bat', '.cmd', '.scr', '.ps1', '.vbs', '.sh', '.msi', '.com', '.pif', '.jar'];
                for (const att of attachments) {
                    const ext = att.filename ? path.extname(att.filename).toLowerCase() : '';
                    if (dangerousExtensions.includes(ext)) {
                        throw ApiError.badRequest(`Attachment "${att.filename}" has a disallowed file extension.`, ErrorCode.INVALID_FILE_TYPE);
                    }
                }
            }

            // Validate recipient email formats - skip invalid ones instead of blocking
            const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
            const invalidEmails = recipients.filter(r => !emailRegex.test(r.email));
            const validRecipients = recipients.filter(r => emailRegex.test(r.email));

            // If ALL emails are invalid, return early
            if (validRecipients.length === 0) {
                res.json({
                    message: `Campaign complete: 0 sent, ${invalidEmails.length} failed (all emails were invalid)`,
                    campaignRunId: null,
                    summary: { total: recipients.length, sent: 0, failed: invalidEmails.length },
                });
                return;
            }

            // Log received attachments
            console.log(`[CampaignController] Received ${attachments?.length || 0} attachments`);

            // Get the user's email credential
            const credentials = await credentialService.getCredentials(userId);

            if (!credentials || credentials.length === 0) {
                throw ApiError.badRequest(
                    'No email credentials found. Please go back to Step 1 and add your Gmail credentials.',
                    ErrorCode.MISSING_REQUIRED_FIELD
                );
            }

            // Find the credential by email or use default
            let credential;
            if (credentialEmail) {
                const cred = credentials.find(c => c.email === credentialEmail);
                if (!cred) {
                    const availableEmails = credentials.map(c => c.email).join(', ');
                    throw ApiError.notFound(
                        `Sender email "${credentialEmail}" not found in your saved credentials. Available credentials: ${availableEmails}. Please go back to Step 1 and add this email.`
                    );
                }
                credential = await credentialService.getCredentialById(userId, cred.id);
            } else {
                const defaultCred = credentials.find(c => c.isDefault) || credentials[0];
                if (!defaultCred) {
                    throw ApiError.badRequest(
                        'No email credential configured. Please go back to Step 1 and set up your sender email.',
                        ErrorCode.MISSING_REQUIRED_FIELD
                    );
                }
                credential = await credentialService.getCredentialById(userId, defaultCred.id);
            }

            // Create a campaign run for tracking progress
            const allEmails = recipients.map(r => r.email);
            const run = campaignRunManager.createRun(userId, allEmails);
            const runId = run.id;

            // Mark invalid emails as failed immediately
            invalidEmails.forEach(r => {
                campaignRunManager.updateRecipient(runId, r.email, 'failed', `Invalid email format: "${r.email}". Skipped.`);
                trackingService.recordDelivery({
                    userId,
                    recipientEmail: r.email,
                    subject: cleanSubject,
                    campaignId: runId,
                    status: 'failed',
                    errorMessage: `Invalid email format: "${r.email}". Skipped.`,
                });
                // Emit to WebSocket room
                io.to(`campaign:${runId}`).emit('email:status', {
                    email: r.email,
                    status: 'failed',
                    error: `Invalid email format: "${r.email}". Skipped.`,
                });
            });

            // Respond immediately with the run ID
            res.json({
                message: 'Campaign started',
                campaignRunId: runId,
                summary: { total: recipients.length, validCount: validRecipients.length, invalidCount: invalidEmails.length },
            });

            // ====== BACKGROUND: Send emails asynchronously ======
            // This runs AFTER the response has been sent to the client
            setImmediate(async () => {
                const credentialData = { email: credential.email, appPassword: credential.appPassword };

                for (let i = 0; i < validRecipients.length; i++) {
                    const recipient = validRecipients[i];

                    // Mark as sending
                    campaignRunManager.updateRecipient(runId, recipient.email, 'sending');
                    io.to(`campaign:${runId}`).emit('email:status', {
                        email: recipient.email,
                        status: 'sending',
                    });

                    let personalizedSubject = cleanSubject;

                    try {
                        // Personalize subject and body
                        personalizedSubject = emailService.personalizeContent(cleanSubject, recipient).replace(/[\r\n]+/g, ' ').trim();
                        const personalizedBody = emailService.personalizeContent(body, recipient);

                        // Prepare clean HTML email (no tracking pixels, no redirect links)
                        const { html: emailHtml } = await emailService.prepareTrackedEmail({
                            recipientEmail: recipient.email,
                            subject: personalizedSubject,
                            html: personalizedBody,
                            campaignId: runId,
                        });

                        // Send email with attachments
                        const emailAttachments = attachments?.map(att => ({
                            filename: att.filename,
                            content: Buffer.from(att.content, 'base64'),
                        }));

                        await emailService.sendEmail(
                            credentialData,
                            {
                                from: credential.email,
                                to: recipient.email,
                                subject: personalizedSubject,
                                html: emailHtml,
                                attachments: emailAttachments,
                            }
                        );

                        // Record delivery success in delivery log
                        await trackingService.recordDelivery({
                            userId,
                            recipientEmail: recipient.email,
                            subject: personalizedSubject,
                            campaignId: runId,
                            status: 'sent',
                        });

                        // Mark as sent
                        campaignRunManager.updateRecipient(runId, recipient.email, 'sent');
                        io.to(`campaign:${runId}`).emit('email:status', {
                            email: recipient.email,
                            status: 'sent',
                        });
                    } catch (error: any) {
                        console.error(`Failed to send to ${recipient.email}:`, error.message);

                        // User-friendly error messages
                        let friendlyError = error.message || 'Unknown error';
                        if (error.message?.includes('Invalid login') || error.message?.includes('authentication failed')) {
                            friendlyError = 'Gmail authentication failed. Please check your App Password in Step 1.';
                        } else if (error.message?.includes('Recipient address rejected') || error.message?.includes('550')) {
                            friendlyError = `Recipient email "${recipient.email}" does not exist or was rejected by the mail server.`;
                        } else if (error.message?.includes('timeout') || error.message?.includes('ETIMEDOUT')) {
                            friendlyError = 'Connection timed out. Gmail may be temporarily unavailable.';
                        } else if (error.message?.includes('ECONNREFUSED') || error.message?.includes('ESOCKET')) {
                            friendlyError = 'Could not connect to Gmail servers. Please check your internet connection.';
                        } else if (error.message?.includes('Rate limit') || error.message?.includes('too many')) {
                            friendlyError = 'Gmail rate limit reached. Please wait a few minutes and try again with fewer recipients.';
                        }

                        // Record delivery failure in delivery log
                        await trackingService.recordDelivery({
                            userId,
                            recipientEmail: recipient.email,
                            subject: personalizedSubject,
                            campaignId: runId,
                            status: 'failed',
                            errorMessage: friendlyError,
                        });

                        // Mark as failed
                        campaignRunManager.updateRecipient(runId, recipient.email, 'failed', friendlyError);
                        io.to(`campaign:${runId}`).emit('email:status', {
                            email: recipient.email,
                            status: 'failed',
                            error: friendlyError,
                        });
                    }

                    // Batch delay - wait between batches (with smart jitter if enabled)
                    if ((i + 1) % batchSize === 0 && i < validRecipients.length - 1) {
                        const jitterFactor = useJitter ? (0.8 + Math.random() * 0.4) : 1.0;
                        const waitMs = Math.round(batchDelay * 1000 * jitterFactor);
                        await new Promise(resolve => setTimeout(resolve, waitMs));
                    } else if (i < validRecipients.length - 1) {
                        // Delay between individual emails (with natural jitter if enabled)
                        const emailDelay = useJitter ? Math.floor(300 + Math.random() * 500) : 300;
                        await new Promise(resolve => setTimeout(resolve, emailDelay));
                    }
                }

                // Emit completion
                const finalRun = campaignRunManager.getRun(runId);
                io.to(`campaign:${runId}`).emit('campaign:completed', {
                    runId,
                    sentCount: finalRun?.sentCount || 0,
                    failedCount: finalRun?.failedCount || 0,
                    totalCount: finalRun?.totalCount || 0,
                });

                console.log(`[CampaignController] Campaign run ${runId} completed: ${finalRun?.sentCount} sent, ${finalRun?.failedCount} failed`);
            });

        } catch (error) {
            next(error);
        }
    }

    /**
     * POST /api/campaigns/validate-domains
     * Verify MX DNS records for recipient domains
     */
    async validateDomains(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const { emails } = req.body;
            if (!emails || !Array.isArray(emails)) {
                throw ApiError.badRequest('Array of recipient emails is required', ErrorCode.MISSING_REQUIRED_FIELD);
            }

            const results = await emailValidatorService.validateEmailsMx(emails);
            res.json({ results });
        } catch (error) {
            next(error);
        }
    }

    /**
     * GET /api/campaigns/run/:runId/status
     * Polling fallback for campaign run progress
     */
    async getCampaignRunStatus(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const { runId } = req.params;
            const status = campaignRunManager.getRunStatus(runId);

            if (!status) {
                throw ApiError.notFound('Campaign run not found or expired.');
            }

            // Verify the run belongs to this user
            const run = campaignRunManager.getRun(runId);
            if (run && run.userId !== req.user!.id) {
                throw ApiError.notFound('Campaign run not found.');
            }

            res.json({ ...status });
        } catch (error) {
            next(error);
        }
    }
}

export default new CampaignController();

