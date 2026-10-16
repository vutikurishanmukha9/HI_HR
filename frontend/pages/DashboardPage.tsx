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
import AnalyticsDashboard from '../components/analytics/AnalyticsDashboard';
import { useToast } from '../components/Toast';

/**
 * DashboardPage Component - Main Campaign Workflow
 * Uses useCampaign hook for state management
 */
const DashboardPage: React.FC = () => {
    const { user, logout } = useAuth();
    const campaign = useCampaign();
    const toast = useToast();

    // Tab state from URL
    const getInitialTab = () => {
        const searchParams = new URLSearchParams(window.location.search);
        return searchParams.get('tab') === 'analytics' ? 'analytics' : 'campaign';
    };
    
    const [activeTab, setActiveTab] = useState<'campaign' | 'analytics'>(getInitialTab);

    // Update URL when tab changes
    useEffect(() => {
        const url = new URL(window.location.href);
        if (activeTab === 'analytics') {
            url.searchParams.set('tab', 'analytics');
        } else {
            url.searchParams.delete('tab');
        }
        window.history.replaceState({}, '', url.toString());
    }, [activeTab]);

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

    // Show success toast when campaign finishes
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
                toast.error('Authentication Error', result.error);
            }
            return result;
        } else {
            toast.success('Credentials Saved', 'Your email credentials have been saved securely.');
            return result;
        }
    };

    return (
        <div className="min-h-screen py-8 px-4 sm:px-6 bg-canvas">
            {/* Confetti celebration on completion */}
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

            <div className="max-w-6xl mx-auto space-y-6">
                {/* Modern Top Header Bar */}
                <header className="surface-card p-4 sm:p-5 flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                    <div className="flex items-center gap-3">
                        <div className="w-10 h-10 rounded-none bg-accent text-white flex items-center justify-center font-bold text-base shadow-xs">
                            Hi
                        </div>
                        <div>
                            <div className="flex items-center gap-2">
                                <span className="font-bold text-lg text-ink tracking-tight">HiHR</span>
                                <span className="px-2 py-0.5 text-[11px] font-medium rounded-none bg-accent-light text-accent border border-neutral-200">
                                    Outreach Platform
                                </span>
                            </div>
                            <p className="text-xs text-ink-muted">High-deliverability recruitment campaigns</p>
                        </div>
                    </div>

                    <div className="flex items-center gap-3 w-full sm:w-auto justify-between sm:justify-end border-t sm:border-t-0 pt-3 sm:pt-0 border-border">
                        <div className="flex items-center gap-2 min-w-0">
                            <div className="w-8 h-8 rounded-none bg-neutral-100 border border-border flex items-center justify-center text-xs font-semibold text-ink-secondary flex-shrink-0">
                                {user?.email ? user.email.charAt(0).toUpperCase() : 'U'}
                            </div>
                            <div className="text-left min-w-0">
                                <p className="text-[11px] font-semibold uppercase tracking-wider text-ink-muted">Signed in</p>
                                <p className="text-xs font-medium text-ink truncate max-w-[160px] sm:max-w-[200px]" title={user?.email}>
                                    {user?.email}
                                </p>
                            </div>
                        </div>
                        <button
                            onClick={logout}
                            className="btn btn-secondary text-xs px-3 py-1.5 ml-2 flex-shrink-0"
                        >
                            Sign Out
                        </button>
                    </div>
                </header>

                {/* Tab Switcher */}
                <div className="flex justify-center">
                    <div className="p-1 rounded-none bg-card border border-border shadow-xs inline-flex gap-1">
                        <button
                            onClick={() => setActiveTab('campaign')}
                            className={`flex items-center gap-2 px-5 py-2 rounded-none text-sm font-semibold transition-all duration-150 ${
                                activeTab === 'campaign'
                                    ? 'bg-accent text-white shadow-xs'
                                    : 'text-ink-secondary hover:text-ink hover:bg-neutral-100'
                            }`}
                        >
                            <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                                <path strokeLinecap="round" strokeLinejoin="round" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" />
                            </svg>
                            Campaign Builder
                        </button>
                        <button
                            onClick={() => setActiveTab('analytics')}
                            className={`flex items-center gap-2 px-5 py-2 rounded-none text-sm font-semibold transition-all duration-150 ${
                                activeTab === 'analytics'
                                    ? 'bg-accent text-white shadow-xs'
                                    : 'text-ink-secondary hover:text-ink hover:bg-neutral-100'
                            }`}
                        >
                            <svg className="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                                <path strokeLinecap="round" strokeLinejoin="round" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
                            </svg>
                            Analytics
                        </button>
                    </div>
                </div>

                {/* Dashboard Stats - Show when there's campaign activity */}
                {activeTab === 'campaign' && (totalSent > 0 || campaign.recipients.length > 0) && (
                    <DashboardStats
                        totalSent={totalSent}
                        successRate={successRate}
                        todaySent={totalSent}
                        pendingCount={campaign.recipients.length - totalSent - totalFailed}
                        onReset={campaign.handleReset}
                    />
                )}

                {/* Main Content Area */}
                {activeTab === 'analytics' ? (
                    <AnalyticsDashboard />
                ) : (
                    <div className="space-y-6">
                        {/* Step Indicator */}
                        <StepIndicator currentStep={campaign.step} totalSteps={campaign.totalSteps} />

                        {/* Main Step Card */}
                        <div className="surface-card p-5 sm:p-8">
                            <div className="animate-fadeIn">
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
                                        availableTags={campaign.availableTags}
                                        initialTemplate={campaign.emailTemplate}
                                    />
                                )}

                                {campaign.step === 4 && campaign.credentials && (
                                    <>
                                        {/* Preview Button */}
                                        <div className="flex justify-end mb-4">
                                            <button
                                                onClick={() => setShowPreview(true)}
                                                className="btn btn-secondary text-xs flex items-center gap-2 px-3 py-1.5"
                                            >
                                                <svg className="w-4 h-4 text-ink-secondary" fill="none" stroke="currentColor" viewBox="0 0 24 24">
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
                                                        toast.warning(
                                                            'Partially Sent',
                                                            `${r.sent} sent, ${r.failed} failed. Hover over failed recipients to see details.`
                                                        );
                                                    } else if (r.failed > 0 && r.sent === 0) {
                                                        const firstError = r.failedDetails?.[0]?.error || 'Unknown error';
                                                        toast.error('All Emails Failed', firstError);
                                                    }
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
                    </div>
                )}

                {/* Modern Minimal Footer */}
                <footer className="pt-4 pb-8 text-center text-xs text-ink-muted border-t border-border space-y-1">
                    <p>
                        HiHR Outreach Engine · End-to-end encrypted SMTP credentials
                    </p>
                    <div className="flex items-center justify-center gap-1.5 text-[11px] text-ink-muted">
                        <span className="w-2 h-2 rounded-xs bg-emerald-500 animate-pulse"></span>
                        <span>Backend active & connected</span>
                    </div>
                </footer>
            </div>
        </div>
    );
};

export default DashboardPage;
