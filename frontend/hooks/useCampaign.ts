import { useState, useCallback } from 'react';
import { Credentials, Recipient, EmailTemplate, EmailStatus, SendProgressState } from '../types';
import apiClient from '../services/api';

/**
 * Hook for managing campaign workflow state
 * Extracts complex state logic from App component
 */
export function useCampaign() {
    const [step, setStep] = useState(1);
    const [credentials, setCredentials] = useState<Credentials | null>(null);
    const [recipients, setRecipients] = useState<Recipient[]>([]);
    const [emailTemplate, setEmailTemplate] = useState<EmailTemplate>({
        subject: '',
        body: '',
        attachments: [],
    });

    // Send progress state
    const [sendProgress, setSendProgress] = useState<Record<string, SendProgressState>>({});
    const [isSending, setIsSending] = useState(false);
    const [isCampaignFinished, setIsCampaignFinished] = useState(false);
    const [scheduledTime, setScheduledTime] = useState<Date | null>(null);

    const totalSteps = 4;

    // Handle saving credentials
    const handleCredentialsSave = useCallback(async (creds: Credentials) => {
        try {
            await apiClient.saveCredential(creds.email, creds.appPassword, true);
            setCredentials(creds);
            setStep(2);
            return { success: true };
        } catch (error: any) {
            let errorMessage = 'Failed to save credentials';

            if (error instanceof Error) {
                errorMessage = error.message;
            } else if (typeof error === 'string') {
                errorMessage = error;
            } else if (error && typeof error === 'object') {
                errorMessage = error.message || error.error || JSON.stringify(error);
            }

            // Handle "already exists" - not really an error
            if (errorMessage.toLowerCase().includes('already exists')) {
                setCredentials(creds);
                setStep(2);
                return { success: true };
            }

            return { success: false, error: errorMessage };
        }
    }, []);

    // Handle recipients upload
    const handleRecipientsUpload = useCallback((uploadedRecipients: Recipient[]) => {
        setRecipients(uploadedRecipients);
        setStep(3);
    }, []);

    // Handle email compose
    const handleEmailCompose = useCallback((template: EmailTemplate) => {
        setEmailTemplate(template);
        setStep(4);
    }, []);

    // Handle navigation
    const handleBack = useCallback(() => {
        if (step > 1) setStep(step - 1);
    }, [step]);

    // Handle sending campaign
    const handleSendCampaign = useCallback(async (config: {
        time?: Date;
        recipientsToSend: Recipient[];
        batchSize: number;
        batchDelay: number;
    }) => {
        if (!credentials) return;

        // Handle scheduling
        if (config.time) {
            setScheduledTime(config.time);
            return;
        }

        setIsSending(true);

        // Initialize progress for all recipients
        const initialProgress: Record<string, SendProgressState> = {};
        config.recipientsToSend.forEach(r => {
            initialProgress[r.email] = { status: EmailStatus.Queued };
        });
        setSendProgress(initialProgress);

        try {
            // Mark all as sending
            config.recipientsToSend.forEach(r => {
                setSendProgress(prev => ({
                    ...prev,
                    [r.email]: { status: EmailStatus.Sending }
                }));
            });

            // Call backend API
            const result = await apiClient.sendCampaign({
                credentialEmail: credentials.email,
                subject: emailTemplate.subject,
                body: emailTemplate.body,
                recipients: config.recipientsToSend,
                attachments: emailTemplate.attachments,
                batchSize: config.batchSize,
                batchDelay: config.batchDelay,
            });

            // Update progress based on results
            result.results.forEach(r => {
                setSendProgress(prev => ({
                    ...prev,
                    [r.email]: {
                        status: r.status === 'sent' ? EmailStatus.Sent : EmailStatus.Failed,
                        error: r.error,
                    }
                }));
            });

            setIsCampaignFinished(true);

            // Return summary for toast notifications
            const sentCount = result.results.filter(r => r.status === 'sent').length;
            const failedCount = result.results.filter(r => r.status === 'failed').length;
            const failedDetails = result.results
                .filter(r => r.status === 'failed')
                .map(r => ({ email: r.email, error: r.error || 'Unknown error' }));

            return {
                success: true,
                sent: sentCount,
                failed: failedCount,
                failedDetails,
            };
        } catch (error: any) {
            console.error('Campaign send error:', error);

            // Parse error message for user-friendly display
            let errorMessage = 'Campaign failed';
            let errorType: 'credential' | 'validation' | 'network' | 'unknown' = 'unknown';

            if (error instanceof Error) {
                errorMessage = error.message;
            } else if (typeof error === 'string') {
                errorMessage = error;
            } else if (error && typeof error === 'object') {
                errorMessage = error.message || error.error || JSON.stringify(error);
            }

            // Categorize the error
            const msg = errorMessage.toLowerCase();
            if (msg.includes('credential') || msg.includes('not found in your saved') || msg.includes('app password') || msg.includes('authentication')) {
                errorType = 'credential';
            } else if (msg.includes('invalid email') || msg.includes('no recipients') || msg.includes('subject and body')) {
                errorType = 'validation';
            } else if (msg.includes('network') || msg.includes('fetch') || msg.includes('connect') || msg.includes('timeout')) {
                errorType = 'network';
            }

            // Mark all as failed
            config.recipientsToSend.forEach(r => {
                setSendProgress(prev => ({
                    ...prev,
                    [r.email]: {
                        status: EmailStatus.Failed,
                        error: errorMessage
                    }
                }));
            });
            setIsCampaignFinished(true);

            return {
                success: false,
                error: errorMessage,
                errorType,
            };
        } finally {
            setIsSending(false);
        }
    }, [credentials, emailTemplate]);

    // Handle cancel schedule
    const handleCancelSchedule = useCallback(() => {
        setScheduledTime(null);
    }, []);

    // Reset campaign
    const handleReset = useCallback(() => {
        setStep(1);
        setCredentials(null);
        setRecipients([]);
        setEmailTemplate({ subject: '', body: '', attachments: [] });
        setSendProgress({});
        setIsSending(false);
        setIsCampaignFinished(false);
        setScheduledTime(null);
    }, []);

    return {
        // State
        step,
        credentials,
        recipients,
        emailTemplate,
        sendProgress,
        isSending,
        isCampaignFinished,
        scheduledTime,
        totalSteps,

        // Actions
        handleCredentialsSave,
        handleRecipientsUpload,
        handleEmailCompose,
        handleBack,
        handleSendCampaign,
        handleCancelSchedule,
        handleReset,
    };
}

export default useCampaign;
