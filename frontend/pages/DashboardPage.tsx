import React, { useState, useEffect } from 'react';
import { useAuth } from '../contexts/AuthContext';
import useCampaign from '../hooks/useCampaign';
import StepIndicator from '../components/StepIndicator';
import CredentialsForm from '../components/CredentialsForm';
import RecipientUploader from '../components/RecipientUploader';
import EmailComposer from '../components/EmailComposer';
import ReviewAndSend from '../components/ReviewAndSend';
import DashboardStats from '../components/DashboardStats';
import Confetti from '../components/Confetti';
import EmailPreview from '../components/EmailPreview';
import { useToast } from '../components/Toast';

/**
 * DashboardPage Component - Main Campaign Workflow
 * Uses useCampaign hook for state management
 */
const DashboardPage: React.FC = () => {
    const { user, logout } = useAuth();
    const campaign = useCampaign();
    const toast = useToast();

    // UI state
    const [showConfetti, setShowConfetti] = useState(false);
    const [showPreview, setShowPreview] = useState(false);
    const [prevStep, setPrevStep] = useState(campaign.step);

    // Track step changes for animation
    useEffect(() => {
        if (campaign.step !== prevStep) {
            setPrevStep(campaign.step);
        }
    }, [campaign.step, prevStep]);

    // Track if we've shown the celebration (prevent multiple toasts)
    const celebrationShownRef = React.useRef(false);

    // Show confetti when campaign finishes successfully
    useEffect(() => {
        if (campaign.isCampaignFinished && !celebrationShownRef.current) {
            const sentCount = Object.values(campaign.sendProgress).filter(
                p => p.status === 'Sent'
            ).length;
            if (sentCount > 0) {
                celebrationShownRef.current = true;
                setShowConfetti(true);
                toast.success('Campaign Complete!', `Successfully sent ${sentCount} emails`);
            }
        }
        // Reset when campaign is reset
        if (!campaign.isCampaignFinished) {
            celebrationShownRef.current = false;
        }
    }, [campaign.isCampaignFinished, campaign.sendProgress, toast]);

    // Calculate stats
    const totalSent = Object.values(campaign.sendProgress).filter(p => p.status === 'Sent').length;
    const totalFailed = Object.values(campaign.sendProgress).filter(p => p.status === 'Failed').length;
    const successRate = totalSent + totalFailed > 0
        ? Math.round((totalSent / (totalSent + totalFailed)) * 100)
        : 100;

    const handleCredentialsSave = async (creds: { email: string; appPassword: string }) => {
        const result = await campaign.handleCredentialsSave(creds);
        if (!result.success && result.error) {
            if (result.error.toLowerCase().includes('network') || result.error.toLowerCase().includes('fetch')) {
                toast.error('Connection Failed', 'Unable to connect to the server.');
            } else if (result.error.toLowerCase().includes('unauthorized') || result.error.toLowerCase().includes('401')) {
                toast.error('Session Expired', 'Please log in again.');
            } else {
                toast.error('Error', result.error);
            }
        } else {
            toast.success('Credentials Saved', 'Your email credentials have been saved securely.');
        }
    };

    return (
        <div className="min-h-screen py-8 px-4">
            {/* Confetti Animation */}
            <Confetti show={showConfetti} onComplete={() => setShowConfetti(false)} />

            {/* Email Preview Modal */}
            {campaign.credentials && (
                <EmailPreview
                    isOpen={showPreview}
                    onClose={() => setShowPreview(false)}
                    emailTemplate={campaign.emailTemplate}
                    recipients={campaign.recipients}
                    senderEmail={campaign.credentials.email}
                />
            )}

            <div className="max-w-4xl mx-auto">
                {/* Premium Header */}
                <div className="glass-card p-5 sm:p-6 mb-6">
                    <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                        <div className="flex items-center gap-3">
                            <div className="w-11 h-11 rounded-xl flex items-center justify-center shadow-lg"
                                style={{ background: 'linear-gradient(135deg, #6366f1 0%, #a855f7 50%, #06b6d4 100%)' }}>
                                <span className="text-white font-extrabold text-lg" style={{ fontFamily: "'Outfit', sans-serif" }}>Hi</span>
                            </div>
                            <div>
                                <h1 className="text-2xl font-extrabold text-brand">HiHR</h1>
                                <p className="text-xs" style={{ color: '#64748b' }}>Smart HR Email Outreach</p>
                            </div>
                        </div>
                        <div className="flex items-center gap-3 w-full sm:w-auto">
                            <div className="text-right flex-1 sm:flex-none">
                                <p className="text-[10px] uppercase tracking-wider" style={{ color: '#64748b' }}>Logged in as</p>
                                <p className="text-sm font-medium truncate max-w-[150px] sm:max-w-none" style={{ color: '#f1f5f9' }}>{user?.email}</p>
                            </div>
                            <button
                                onClick={logout}
                                className="px-4 py-2 rounded-lg text-sm font-medium transition-all duration-300"
                                style={{ border: '1px solid rgba(148, 163, 184, 0.15)', color: '#94a3b8' }}
                            >
                                Logout
                            </button>
                        </div>
                    </div>
                </div>

                {/* Dashboard Stats - Show when there's activity */}
                {(totalSent > 0 || campaign.recipients.length > 0) && (
                    <DashboardStats
                        totalSent={totalSent}
                        successRate={successRate}
                        todaySent={totalSent}
                        pendingCount={campaign.recipients.length - totalSent - totalFailed}
                    />
                )}

                {/* Step Indicator */}
                <StepIndicator currentStep={campaign.step} totalSteps={campaign.totalSteps} />

                {/* Main Content with Step Transition */}
                <div className="glass-card p-4 sm:p-8 mt-6">
                    <div className="step-enter">
                        {campaign.step === 1 && (
                            <CredentialsForm
                                onSave={handleCredentialsSave}
                                initialCredentials={campaign.credentials}
                            />
                        )}

                        {campaign.step === 2 && (
                            <RecipientUploader
                                onUpload={campaign.handleRecipientsUpload}
                                onBack={campaign.handleBack}
                            />
                        )}

                        {campaign.step === 3 && (
                            <EmailComposer
                                onCompose={campaign.handleEmailCompose}
                                onBack={campaign.handleBack}
                                recipients={campaign.recipients}
                                initialTemplate={campaign.emailTemplate}
                            />
                        )}

                        {campaign.step === 4 && campaign.credentials && (
                            <>
                                {/* Preview Button */}
                                <div className="flex justify-end mb-4">
                                    <button
                                        onClick={() => setShowPreview(true)}
                                        className="flex items-center gap-2 px-4 py-2 rounded-lg bg-white/10 text-white hover:bg-white/20 transition-colors"
                                    >
                                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
                                        </svg>
                                        Preview Email
                                    </button>
                                </div>

                                <ReviewAndSend
                                    credentials={campaign.credentials}
                                    recipients={campaign.recipients}
                                    emailTemplate={campaign.emailTemplate}
                                    sendProgress={campaign.sendProgress}
                                    isSending={campaign.isSending}
                                    isCampaignFinished={campaign.isCampaignFinished}
                                    scheduledTime={campaign.scheduledTime}
                                    onScheduleOrSend={async (config) => {
                                        const result = await campaign.handleSendCampaign(config);
                                        if (!result) return;

                                        if (!result.success) {
                                            // Whole campaign failed
                                            const errorType = (result as any).errorType;
                                            const errorMsg = (result as any).error || 'Campaign failed';

                                            if (errorType === 'credential') {
                                                toast.error('Credential Error', errorMsg);
                                            } else if (errorType === 'validation') {
                                                toast.error('Validation Error', errorMsg);
                                            } else if (errorType === 'network') {
                                                toast.error('Connection Error', errorMsg);
                                            } else {
                                                toast.error('Send Failed', errorMsg);
                                            }
                                        } else if (result.success) {
                                            const r = result as any;
                                            if (r.failed > 0 && r.sent > 0) {
                                                // Partial success
                                                toast.warning(
                                                    'Partially Sent',
                                                    `${r.sent} sent, ${r.failed} failed. Hover over failed recipients to see details.`
                                                );
                                            } else if (r.failed > 0 && r.sent === 0) {
                                                // All failed
                                                const firstError = r.failedDetails?.[0]?.error || 'Unknown error';
                                                toast.error('All Emails Failed', firstError);
                                            }
                                            // Full success is handled by the confetti/celebration effect
                                        }
                                    }}
                                    onCancelSchedule={campaign.handleCancelSchedule}
                                    onBack={campaign.handleBack}
                                    onReset={campaign.handleReset}
                                />
                            </>
                        )}
                    </div>
                </div>

                {/* Footer */}
                <div className="text-center mt-8 text-xs" style={{ color: '#475569' }}>
                    <p>
                        Powered by <span className="font-semibold" style={{ color: '#94a3b8' }}>HiHR</span> · Credentials encrypted end-to-end
                    </p>
                    <p className="mt-1 flex items-center justify-center gap-1">
                        <span className="w-1.5 h-1.5 rounded-full" style={{ background: '#14b8a6' }}></span>
                        <span>Backend connected</span>
                    </p>
                </div>
            </div>
        </div>
    );
};

export default DashboardPage;
