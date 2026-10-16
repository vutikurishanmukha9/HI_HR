import { useState, useCallback, useEffect, useRef, useMemo } from 'react';
import { Credentials, Recipient, EmailTemplate, EmailStatus, SendProgressState } from '../types';
import apiClient from '../services/api';
import { io, Socket } from 'socket.io-client';

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:5000';
/**
 * Hook for managing campaign workflow state
 * Extracts complex state logic from App component
 */

const mapStatusToEnum = (status: string): EmailStatus => {
    switch (status.toLowerCase()) {
        case 'queued': return EmailStatus.Queued;
        case 'sending': return EmailStatus.Sending;
        case 'sent': return EmailStatus.Sent;
        case 'failed': return EmailStatus.Failed;
        default: return EmailStatus.Queued;
    }
};

export function useCampaign() {
    const [step, setStep] = useState(1);
    const [credentials, setCredentials] = useState<Credentials | null>(null);
    const [recipients, setRecipients] = useState<Recipient[]>([]);
    
    // Compute available tags dynamically from all recipients
    const availableTags = useMemo(() => {
        if (!recipients || recipients.length === 0) {
            return ['name', 'company', 'role'];
        }
        const tagSet = new Set<string>();
        recipients.forEach(r => {
            Object.keys(r).forEach(key => {
                if (key !== 'email' && !key.startsWith('_') && key !== 'id') {
                    tagSet.add(key);
                }
            });
        });
        return Array.from(tagSet);
    }, [recipients]);
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

    // WebSocket & Polling refs
    const socketRef = useRef<Socket | null>(null);
    const pollIntervalRef = useRef<NodeJS.Timeout | null>(null);

    const totalSteps = 4;

    // Cleanup on unmount
    useEffect(() => {
        return () => {
            if (socketRef.current) {
                socketRef.current.disconnect();
            }
            if (pollIntervalRef.current) {
                clearInterval(pollIntervalRef.current);
            }
        };
    }, []);

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
        time?: Date | null;
        recipientsToSend: Recipient[];
        batchSize: number;
        batchDelay: number;
        useJitter?: boolean;
    }) => {
        if (!credentials) return;

        if (config.time) {
            setScheduledTime(config.time);
            return;
        }

        setIsSending(true);

        const initialProgress: Record<string, SendProgressState> = {};
        config.recipientsToSend.forEach(r => {
            initialProgress[r.email] = { status: EmailStatus.Queued };
        });
        setSendProgress(initialProgress);

        try {
            // 1. Call Backend (Fire and Forget)
            const result = await apiClient.sendCampaign({
                credentialEmail: credentials.email,
                subject: emailTemplate.subject,
                body: emailTemplate.body,
                recipients: config.recipientsToSend,
                attachments: emailTemplate.attachments,
                batchSize: config.batchSize,
                batchDelay: config.batchDelay,
                useJitter: config.useJitter,
            });

            if (!result.campaignRunId) {
                // If it returned no runId, it means all were invalid and it finished immediately
                const finalProgress: Record<string, SendProgressState> = {};
                config.recipientsToSend.forEach(r => {
                    finalProgress[r.email] = { status: EmailStatus.Failed, error: 'Invalid email' };
                });
                setSendProgress(finalProgress);
                setIsSending(false);
                setIsCampaignFinished(true);
                return { success: false, error: 'All emails were invalid', errorType: 'validation' };
            }

            const runId = result.campaignRunId;

            // 2. Setup WebSocket
            socketRef.current = io(API_BASE_URL, {
                transports: ['websocket', 'polling'],
                auth: {
                    token: localStorage.getItem('accessToken') || '',
                },
            });
            
            socketRef.current.on('connect', () => {
                console.log('WebSocket connected, joining campaign run:', runId);
                socketRef.current?.emit('join:campaign', runId);
                
                // Clear any polling since WS is active
                if (pollIntervalRef.current) {
                    clearInterval(pollIntervalRef.current);
                    pollIntervalRef.current = null;
                }
            });

            socketRef.current.on('email:status', (data: { email: string; status: string; error?: string }) => {
                setSendProgress(prev => ({
                    ...prev,
                    [data.email]: {
                        status: mapStatusToEnum(data.status),
                        error: data.error,
                    }
                }));
            });

            socketRef.current.on('campaign:completed', () => {
                console.log('Campaign completed via WebSocket');
                finishCampaign();
            });

            // 3. Fallback Polling Setup
            // If WebSocket disconnects or fails, polling takes over
            socketRef.current.on('disconnect', () => {
                console.log('WebSocket disconnected, falling back to polling...');
                startPolling(runId);
            });

            // Always start polling initially just in case WS fails to connect at all
            startPolling(runId);

            // Finish logic
            const finishCampaign = () => {
                if (socketRef.current) socketRef.current.disconnect();
                if (pollIntervalRef.current) clearInterval(pollIntervalRef.current);
                setIsSending(false);
                setIsCampaignFinished(true);
            };

            // Setup polling function
            function startPolling(campaignRunId: string) {
                if (pollIntervalRef.current) return; // Already polling
                
                pollIntervalRef.current = setInterval(async () => {
                    try {
                        const status = await apiClient.getCampaignRunStatus(campaignRunId);
                        
                        // Update progress from polling data
                        setSendProgress(prev => {
                            const newProgress = { ...prev };
                            Object.entries(status.recipients).forEach(([email, data]) => {
                                newProgress[email] = {
                                    status: mapStatusToEnum(data.status),
                                    error: data.error,
                                };
                            });
                            return newProgress;
                        });

                        if (status.status === 'completed') {
                            console.log('Campaign completed via Polling');
                            finishCampaign();
                        }
                    } catch (err) {
                        console.error('Polling error:', err);
                        // If 404, it might have expired or completed a while ago
                    }
                }, 2000); // Poll every 2 seconds
            }

            return { success: true };

        } catch (error: any) {
            console.error('Campaign start error:', error);
            
            // Revert to initial state
            setSendProgress(initialProgress);
            setIsSending(false);
            
            let errorMessage = 'Campaign failed to start';
            if (error instanceof Error) errorMessage = error.message;
            else if (typeof error === 'string') errorMessage = error;
            else if (error && error.message) errorMessage = error.message;

            return { success: false, error: errorMessage, errorType: 'network' };
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
        availableTags,
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
