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
  fullName: string;
  email: string;
  companyName: string;
  jobTitle?: string;
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
