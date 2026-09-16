/**
 * User credentials for Gmail SMTP authentication
 */
export interface Credentials {
  email: string;
  appPassword: string;
}

/**
 * Recipient information for personalized emails
 */
export interface Recipient {
  email: string;
  [key: string]: any;
}

/**
 * Email template with personalization placeholders
 */
export interface EmailTemplate {
  subject: string;
  body: string;
  attachments?: File[];
}

/**
 * Status of an email in the sending process
 */
export enum EmailStatus {
  Queued = 'Queued',
  Sending = 'Sending',
  Sent = 'Sent',
  Failed = 'Failed',
}

/**
 * Progress state for individual email sending
 */
export interface SendProgressState {
  status: EmailStatus;
  error?: string;
}

/**
 * Delivery statistics overview
 */
export interface TrackingStats {
  totalSent: number;
  sentCount: number;
  failedCount: number;
  successRate: number;
}

/**
 * Detailed delivery record for a single recipient
 */
export interface TrackingDetail {
  recipientEmail: string;
  subject?: string;
  campaignId?: string;
  status: 'sent' | 'failed';
  sentAt?: string;
  errorMessage?: string;
}

/**
 * Pagination metadata
 */
export interface PaginationMeta {
  total: number;
  page: number;
  limit: number;
  totalPages: number;
}

/**
 * Saved email template item
 */
export interface SavedTemplate {
  id: string;
  name: string;
  category: string;
  subject: string;
  body: string;
  isSystem?: boolean;
  createdAt?: string;
  updatedAt?: string;
}

/**
 * Spam and deliverability analysis issue
 */
export interface SpamIssue {
  type: string;
  message: string;
  severity: 'high' | 'medium' | 'info';
}

/**
 * Spam and deliverability analysis report
 */
export interface SpamReport {
  score: number;
  rating: 'Excellent' | 'Good' | 'Needs Improvement';
  triggerWords: string[];
  spamWordsFound?: string[];
  issues: SpamIssue[];
}

/**
 * Domain MX check result
 */
export interface DomainMxResult {
  domain: string;
  hasMx: boolean;
  exchange?: string;
  error?: string;
}

/**
 * Saved email credential item from backend
 */
export interface SavedCredential {
  id: string;
  email: string;
  isDefault: boolean;
  createdAt: string;
}
