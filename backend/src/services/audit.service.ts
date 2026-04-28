/**
 * Audit Service
 * Logs important actions for compliance and debugging
 * 
 * Note: Uses file-based logging by default. Enable database logging
 * by running `npx prisma db push` after stopping the dev server.
 */
import logger from '../utils/logger';
import fs from 'fs';
import path from 'path';

export enum AuditAction {
    EMAIL_SENT = 'email_sent',
    EMAIL_FAILED = 'email_failed',
    CREDENTIAL_ADDED = 'credential_added',
    CREDENTIAL_DELETED = 'credential_deleted',
    CAMPAIGN_CREATED = 'campaign_created',
    CAMPAIGN_SENT = 'campaign_sent',
    LOGIN_SUCCESS = 'login_success',
    LOGIN_FAILED = 'login_failed',
    PASSWORD_CHANGED = 'password_changed',
}

interface AuditLogData {
    userId: string;
    action: AuditAction;
    targetEmail?: string;
    details?: Record<string, any>;
    ipAddress?: string;
}

interface AuditLogEntry extends AuditLogData {
    id: string;
    createdAt: string;
}

// Ensure logs directory exists
const LOGS_DIR = path.join(process.cwd(), 'logs');
const AUDIT_LOG_FILE = path.join(LOGS_DIR, 'audit.jsonl');

if (!fs.existsSync(LOGS_DIR)) {
    fs.mkdirSync(LOGS_DIR, { recursive: true });
}

/**
 * Generate unique ID
 */
function generateId(): string {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Create an audit log entry (file-based)
 */
export async function createAuditLog(data: AuditLogData): Promise<void> {
    try {
        const entry: AuditLogEntry = {
            id: generateId(),
            ...data,
            createdAt: new Date().toISOString(),
        };

        // Append to JSONL file (JSON Lines format)
        const line = JSON.stringify(entry) + '\n';
        fs.appendFileSync(AUDIT_LOG_FILE, line);

        logger.debug(`Audit log created: ${data.action}`, {
            userId: data.userId,
            targetEmail: data.targetEmail,
        });
    } catch (error: any) {
        // Don't let audit logging failures break the main flow
        logger.error('Failed to create audit log', {
            error: error.message,
            action: data.action,
        });
    }
}

/**
 * Get audit logs for a user (file-based)
 */
export async function getAuditLogs(
    userId: string,
    options?: {
        action?: AuditAction;
        limit?: number;
    }
): Promise<{ logs: AuditLogEntry[]; total: number }> {
    try {
        if (!fs.existsSync(AUDIT_LOG_FILE)) {
            return { logs: [], total: 0 };
        }

        const content = fs.readFileSync(AUDIT_LOG_FILE, 'utf-8');
        const lines = content.trim().split('\n').filter(Boolean);

        let logs: AuditLogEntry[] = lines
            .map(line => {
                try {
                    return JSON.parse(line);
                } catch {
                    return null;
                }
            })
            .filter((log): log is AuditLogEntry => log !== null)
            .filter(log => log.userId === userId);

        if (options?.action) {
            logs = logs.filter(log => log.action === options.action);
        }

        // Sort by newest first
        logs.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());

        const total = logs.length;

        if (options?.limit) {
            logs = logs.slice(0, options.limit);
        }

        return { logs, total };
    } catch (error: any) {
        logger.error('Failed to read audit logs', { error: error.message });
        return { logs: [], total: 0 };
    }
}

/**
 * Log email sent action
 */
export async function logEmailSent(
    userId: string,
    recipientEmail: string,
    subject: string,
    ipAddress?: string
): Promise<void> {
    await createAuditLog({
        userId,
        action: AuditAction.EMAIL_SENT,
        targetEmail: recipientEmail,
        details: { subject },
        ipAddress,
    });
}

/**
 * Log email failed action
 */
export async function logEmailFailed(
    userId: string,
    recipientEmail: string,
    error: string,
    ipAddress?: string
): Promise<void> {
    await createAuditLog({
        userId,
        action: AuditAction.EMAIL_FAILED,
        targetEmail: recipientEmail,
        details: { error },
        ipAddress,
    });
}

export default {
    createAuditLog,
    getAuditLogs,
    logEmailSent,
    logEmailFailed,
    AuditAction,
};
