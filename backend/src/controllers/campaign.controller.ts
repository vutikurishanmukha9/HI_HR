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
        fullName: string;
        companyName: string;
        jobTitle?: string;
    }>;
    attachments?: Array<{
        filename: string;
        content: string; // base64 encoded
        contentType: string;
    }>;
    batchSize?: number;
    batchDelay?: number;
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
     * This is the secure method - no credentials exposed to client
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
                batchDelay = 60
            } = req.body as SendCampaignRequest;

            if (!recipients || recipients.length === 0) {
                throw ApiError.badRequest('No recipients provided. Please upload at least one recipient.', ErrorCode.MISSING_REQUIRED_FIELD);
            }

            if (!subject || !body) {
                throw ApiError.badRequest('Subject and body are required. Please compose your email before sending.', ErrorCode.MISSING_REQUIRED_FIELD);
            }

            // Validate recipient email formats - skip invalid ones instead of blocking
            const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
            const invalidEmails = recipients.filter(r => !emailRegex.test(r.email));
            const validRecipients = recipients.filter(r => emailRegex.test(r.email));

            // Pre-populate results for invalid emails (mark as failed immediately)
            const results: Array<{ email: string; status: 'sent' | 'failed'; error?: string }> = [];
            invalidEmails.forEach(r => {
                results.push({
                    email: r.email,
                    status: 'failed',
                    error: `Invalid email format: "${r.email}". Skipped.`,
                });
            });

            // If ALL emails are invalid, return early with results
            if (validRecipients.length === 0) {
                res.json({
                    message: `Campaign complete: 0 sent, ${invalidEmails.length} failed (all emails were invalid)`,
                    results,
                    summary: { total: recipients.length, sent: 0, failed: invalidEmails.length },
                });
                return;
            }

            // Log received attachments
            console.log(`[CampaignController] Received ${attachments?.length || 0} attachments`);
            if (attachments && attachments.length > 0) {
                attachments.forEach((att, i) => {
                    console.log(`[CampaignController] Attachment ${i + 1}: ${att.filename}, content length: ${att.content?.length || 0}`);
                });
            }

            // Get the user's email credential
            const credentials = await credentialService.getCredentials(req.user!.id);

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
                credential = await credentialService.getCredentialById(req.user!.id, cred.id);
            } else {
                const defaultCred = credentials.find(c => c.isDefault);
                if (!defaultCred) {
                    throw ApiError.badRequest(
                        'No default email credential configured. Please go back to Step 1 and set up your sender email.',
                        ErrorCode.MISSING_REQUIRED_FIELD
                    );
                }
                credential = await credentialService.getCredentialById(req.user!.id, defaultCred.id);
            }

            // Send emails in batches (only to valid recipients)
            for (let i = 0; i < validRecipients.length; i++) {
                const recipient = validRecipients[i];

                try {
                    // Personalize subject and body
                    const personalizedSubject = emailService.personalizeContent(subject, {
                        fullName: recipient.fullName,
                        companyName: recipient.companyName,
                        jobTitle: recipient.jobTitle || '',
                    });

                    const personalizedBody = emailService.personalizeContent(body, {
                        fullName: recipient.fullName,
                        companyName: recipient.companyName,
                        jobTitle: recipient.jobTitle || '',
                    });

                    // Prepare email with tracking (pixel + link rewriting)
                    const { html: trackedBody } = await emailService.prepareTrackedEmail({
                        recipientEmail: recipient.email,
                        subject: personalizedSubject,
                        html: personalizedBody,
                    });

                    // Send email with attachments - use proper nodemailer format
                    const emailAttachments = attachments?.map(att => ({
                        filename: att.filename,
                        content: Buffer.from(att.content, 'base64'),
                    }));

                    await emailService.sendEmail(
                        { email: credential.email, appPassword: credential.appPassword },
                        {
                            from: credential.email,
                            to: recipient.email,
                            subject: personalizedSubject,
                            html: trackedBody,
                            attachments: emailAttachments,
                        }
                    );

                    results.push({ email: recipient.email, status: 'sent' });
                } catch (error: any) {
                    console.error(`Failed to send to ${recipient.email}:`, error.message);

                    // Provide user-friendly error messages per recipient
                    let friendlyError = error.message || 'Unknown error';
                    if (error.message?.includes('Invalid login') || error.message?.includes('authentication failed')) {
                        friendlyError = 'Gmail authentication failed. Please check your App Password in Step 1.';
                    } else if (error.message?.includes('Recipient address rejected') || error.message?.includes('550')) {
                        friendlyError = `Recipient email "${recipient.email}" does not exist or was rejected by the mail server.`;
                    } else if (error.message?.includes('timeout') || error.message?.includes('ETIMEDOUT')) {
                        friendlyError = 'Connection timed out. Gmail may be temporarily unavailable. Please try again.';
                    } else if (error.message?.includes('ECONNREFUSED') || error.message?.includes('ESOCKET')) {
                        friendlyError = 'Could not connect to Gmail servers. Please check your internet connection.';
                    } else if (error.message?.includes('Rate limit') || error.message?.includes('too many')) {
                        friendlyError = 'Gmail rate limit reached. Please wait a few minutes and try again with fewer recipients.';
                    }

                    results.push({
                        email: recipient.email,
                        status: 'failed',
                        error: friendlyError
                    });
                }

                // Batch delay - wait between batches
                if ((i + 1) % batchSize === 0 && i < validRecipients.length - 1) {
                    await new Promise(resolve => setTimeout(resolve, batchDelay * 1000));
                } else if (i < validRecipients.length - 1) {
                    // Small delay between individual emails (300ms)
                    await new Promise(resolve => setTimeout(resolve, 300));
                }
            }

            const sentCount = results.filter(r => r.status === 'sent').length;
            const failedCount = results.filter(r => r.status === 'failed').length;

            res.json({
                message: `Campaign sent: ${sentCount} succeeded, ${failedCount} failed`,
                results,
                summary: { total: recipients.length, sent: sentCount, failed: failedCount }
            });
        } catch (error) {
            next(error);
        }
    }
}

export default new CampaignController();

