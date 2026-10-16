import React, { useState, useEffect } from 'react';
import DOMPurify from 'dompurify';
import { Credentials, Recipient, EmailTemplate, EmailStatus, SendProgressState } from '../types';

interface ReviewAndSendProps {
    credentials: Credentials | null;
    recipients: Recipient[];
    emailTemplate: EmailTemplate;
    sendProgress: Record<string, SendProgressState>;
    isSending: boolean;
    isCampaignFinished: boolean;
    scheduledTime: Date | null;
    onScheduleOrSend: (config: { time: Date | null, batchSize: number, batchDelay: number, recipientsToSend: Recipient[], useJitter?: boolean }) => void;
    onCancelSchedule: () => void;
    onBack: () => void;
    onReset: () => void;
}

const StatusIndicator: React.FC<{ progress: SendProgressState }> = ({ progress }) => {
    if (!progress) return null;
    const { status, error } = progress;

    switch (status) {
        case EmailStatus.Queued:
            return (
                <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-none text-[11px] font-medium bg-neutral-100 text-neutral-600 border border-neutral-200 flex-shrink-0 whitespace-nowrap">
                    <span className="w-1.5 h-1.5 rounded-none bg-neutral-400 flex-shrink-0" />
                    <span>Queued</span>
                </span>
            );
        case EmailStatus.Sending:
            return (
                <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-none text-[11px] font-medium bg-neutral-100 text-neutral-800 border border-neutral-300 flex-shrink-0 whitespace-nowrap">
                    <span className="w-1.5 h-1.5 rounded-none bg-neutral-900 animate-pulse flex-shrink-0" />
                    <span>Sending…</span>
                </span>
            );
        case EmailStatus.Sent:
            return (
                <span className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-none text-[11px] font-medium bg-emerald-50 text-emerald-700 border border-emerald-200 flex-shrink-0 whitespace-nowrap">
                    <svg className="w-3.5 h-3.5 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                        <path strokeLinecap="round" strokeLinejoin="round" d="M5 13l4 4L19 7" />
                    </svg>
                    <span>Sent</span>
                </span>
            );
        case EmailStatus.Failed:
            return (
                <span title={error || 'Send failed'} className="inline-flex items-center gap-1.5 px-2 py-0.5 rounded-none text-[11px] font-medium bg-danger-light text-danger border border-red-200 cursor-help flex-shrink-0 whitespace-nowrap">
                    <svg className="w-3.5 h-3.5 flex-shrink-0" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2.5}>
                        <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
                    </svg>
                    <span>Failed</span>
                </span>
            );
        default:
            return null;
    }
};

const ReviewAndSend: React.FC<ReviewAndSendProps> = ({
    credentials, recipients, emailTemplate, sendProgress, isSending, isCampaignFinished,
    scheduledTime, onScheduleOrSend, onCancelSchedule, onBack, onReset
}) => {
    const [selectedEmails, setSelectedEmails] = useState<string[]>(() => recipients.map(r => r.email));
    const [campaignRecipients, setCampaignRecipients] = useState<Recipient[] | null>(null);
    const [isScheduling, setIsScheduling] = useState(false);
    const [scheduleDateTime, setScheduleDateTime] = useState('');
    const [scheduleError, setScheduleError] = useState('');
    const [batchSize, setBatchSize] = useState(10);
    const [batchDelay, setBatchDelay] = useState(60);
    const [useJitter, setUseJitter] = useState(true);

    useEffect(() => {
        setSelectedEmails(recipients.map(r => r.email));
        setCampaignRecipients(null);
    }, [recipients]);

    const recipientsForCampaign = recipients.filter(r => selectedEmails.includes(r.email));
    const displayList = campaignRecipients || recipients;
    const isReady = credentials && recipients.length > 0 && emailTemplate.subject && emailTemplate.body;

    const sentCount = Object.values(sendProgress).filter((s: SendProgressState) => s.status === EmailStatus.Sent).length;
    const failedCount = Object.values(sendProgress).filter((s: SendProgressState) => s.status === EmailStatus.Failed).length;
    const campaignSize = campaignRecipients?.length ?? recipientsForCampaign.length;
    const progressPercentage = campaignSize > 0 ? Math.round(((sentCount + failedCount) / campaignSize) * 100) : 0;

    const getMinDateTime = () => {
        const now = new Date();
        now.setMinutes(now.getMinutes() + 1);
        return now.toISOString().slice(0, 16);
    };

    const handleActionClick = () => {
        const config = {
            time: null as Date | null,
            batchSize: Math.max(1, batchSize),
            batchDelay: Math.max(0, batchDelay),
            recipientsToSend: recipientsForCampaign,
            useJitter
        };
        if (isScheduling) {
            if (!scheduleDateTime) { setScheduleError('Please select a date and time.'); return; }
            const scheduleDate = new Date(scheduleDateTime);
            if (scheduleDate.getTime() <= Date.now()) { setScheduleError('Please select a future time.'); return; }
            setScheduleError('');
            config.time = scheduleDate;
        } else {
            setScheduleError('');
        }
        setCampaignRecipients(recipientsForCampaign);
        onScheduleOrSend(config);
    };

    const handleRecipientSelect = (email: string) => setSelectedEmails(prev => prev.includes(email) ? prev.filter(e => e !== email) : [...prev, email]);
    const handleSelectAll = (e: React.ChangeEvent<HTMLInputElement>) => setSelectedEmails(e.target.checked ? recipients.map(r => r.email) : []);
    const getButtonText = () => isSending ? 'Sending…' : isScheduling ? 'Schedule Campaign' : `Launch Campaign (${recipientsForCampaign.length} Emails)`;

    return (
        <div className="max-w-6xl mx-auto">
            {/* Header */}
            <div className="text-center mb-6">
                <div className="inline-flex items-center justify-center w-12 h-12 rounded-none bg-accent-light border border-neutral-200 text-accent mb-3 shadow-xs">
                    <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                </div>
                <h2 className="text-xl font-bold text-ink">Review & Launch</h2>
                <p className="text-sm text-ink-muted mt-1">
                    {isCampaignFinished ? 'Campaign execution completed!' : scheduledTime ? 'Campaign scheduled for dispatch.' : 'Verify recipients and outreach parameters before sending'}
                </p>
            </div>

            {isReady ? (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    {/* Left Column - Summary & Controls */}
                    <div className="space-y-5">
                        <div className="surface-card p-5 border border-border shadow-xs space-y-4">
                            <h3 className="text-sm font-bold text-ink flex items-center gap-2">
                                <svg className="w-4 h-4 text-accent" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
                                </svg>
                                Campaign Summary
                            </h3>

                            <div className="space-y-2.5 text-xs">
                                <div className="flex justify-between items-center py-1.5 border-b border-border gap-2 min-w-0">
                                    <span className="text-ink-secondary flex-shrink-0">From Account</span>
                                    <span className="font-mono text-ink font-medium truncate text-right min-w-0" title={credentials?.email}>{credentials?.email}</span>
                                </div>
                                <div className="flex justify-between items-center py-1.5 border-b border-border gap-2 min-w-0">
                                    <span className="text-ink-secondary flex-shrink-0">Target Outreach</span>
                                    <span className="font-semibold text-ink truncate text-right min-w-0">{recipientsForCampaign.length} recipients selected</span>
                                </div>
                                <div className="flex justify-between items-center py-1.5 border-b border-border gap-2 min-w-0">
                                    <span className="text-ink-secondary flex-shrink-0">Subject</span>
                                    <span className="font-medium text-ink truncate max-w-[260px] text-right min-w-0" title={emailTemplate.subject}>{emailTemplate.subject}</span>
                                </div>
                                {emailTemplate.attachments && emailTemplate.attachments.length > 0 && (
                                    <div className="flex justify-between items-center py-1.5 border-b border-border gap-2 min-w-0">
                                        <span className="text-ink-secondary flex-shrink-0">Attachments</span>
                                        <span className="text-ink font-medium truncate text-right min-w-0">{emailTemplate.attachments.length} files</span>
                                    </div>
                                )}
                            </div>

                            <div className="pt-2">
                                <p className="text-xs font-semibold uppercase tracking-wider text-ink-secondary mb-1.5">Template Preview</p>
                                <div
                                    className="p-3 rounded-none max-h-28 overflow-y-auto text-xs bg-neutral-50 border border-border text-ink-secondary leading-relaxed break-words"
                                    dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(emailTemplate.body) }}
                                />
                            </div>
                        </div>

                        {/* Completion / Scheduled / Send Options Card */}
                        {isCampaignFinished ? (
                            <div className="surface-card p-5 border border-emerald-200 bg-emerald-50/20 shadow-xs space-y-4">
                                <h4 className="font-bold text-sm text-ink flex items-center gap-2">
                                    <span className="w-2 h-2 rounded-none bg-emerald-500" />
                                    Campaign Execution Summary
                                </h4>
                                <div className="grid grid-cols-2 gap-3">
                                    <div className="text-center p-3.5 rounded-none bg-white border border-emerald-200 shadow-xs">
                                        <p className="text-2xl font-bold text-emerald-600 tabular-nums">{sentCount}</p>
                                        <p className="text-xs font-medium text-ink-muted">Delivered</p>
                                    </div>
                                    <div className="text-center p-3.5 rounded-none bg-white border border-red-200 shadow-xs">
                                        <p className="text-2xl font-bold text-danger tabular-nums">{failedCount}</p>
                                        <p className="text-xs font-medium text-ink-muted">Failed</p>
                                    </div>
                                </div>
                                <button onClick={onReset} className="btn btn-primary w-full py-2.5 text-xs shadow-xs">
                                    Start New Outreach Campaign
                                </button>
                            </div>
                        ) : scheduledTime && !isSending ? (
                            <div className="surface-card p-5 border border-neutral-200 bg-neutral-50 shadow-xs space-y-4">
                                <div className="text-center space-y-1">
                                    <p className="text-xs font-semibold text-ink-secondary uppercase tracking-wider">Scheduled Launch</p>
                                    <p className="text-base font-bold text-ink">{scheduledTime.toLocaleString()}</p>
                                </div>
                                <div className="flex gap-2">
                                    <button
                                        onClick={() => onScheduleOrSend({ time: null, batchSize, batchDelay, recipientsToSend: recipientsForCampaign })}
                                        className="btn btn-primary flex-1 text-xs py-2"
                                    >
                                        Send Immediately
                                    </button>
                                    <button onClick={onCancelSchedule} className="btn btn-secondary flex-1 text-xs py-2">
                                        Cancel Schedule
                                    </button>
                                </div>
                            </div>
                        ) : (
                            <div className="surface-card p-5 border border-border shadow-xs space-y-4">
                                <h4 className="font-bold text-sm text-ink">Throttling & Rate Control</h4>
                                <div className="grid grid-cols-2 gap-3">
                                    <div className="space-y-1">
                                        <label className="block text-[11px] font-semibold uppercase tracking-wider text-ink-secondary">
                                            Batch Size
                                        </label>
                                        <input
                                            type="number"
                                            value={batchSize}
                                            onChange={(e) => setBatchSize(parseInt(e.target.value, 10) || 1)}
                                            min="1"
                                            disabled={isSending}
                                            className="input py-1.5 text-xs font-mono"
                                        />
                                    </div>
                                    <div className="space-y-1">
                                        <label className="block text-[11px] font-semibold uppercase tracking-wider text-ink-secondary">
                                            Delay (seconds)
                                        </label>
                                        <input
                                            type="number"
                                            value={batchDelay}
                                            onChange={(e) => setBatchDelay(parseInt(e.target.value, 10) || 0)}
                                            min="0"
                                            disabled={isSending}
                                            className="input py-1.5 text-xs font-mono"
                                        />
                                    </div>
                                </div>

                                <div className="p-3 rounded-none bg-neutral-50 border border-border space-y-1.5">
                                    <div className="flex items-center justify-between">
                                        <label htmlFor="jitter-toggle" className="text-xs font-semibold text-ink cursor-pointer flex items-center gap-1.5">
                                            <svg className="w-4 h-4 text-accent" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 10V3L4 14h7v7l9-11h-7z" />
                                            </svg>
                                            Human-like Smart Jitter (Anti-Bot)
                                        </label>
                                        <input
                                            id="jitter-toggle"
                                            type="checkbox"
                                            checked={useJitter}
                                            onChange={(e) => setUseJitter(e.target.checked)}
                                            disabled={isSending}
                                            className="w-4 h-4 rounded-none text-accent border-border"
                                        />
                                    </div>
                                    <p className="text-[11px] text-ink-muted leading-relaxed">
                                         Varies delay between batches (+/-20%) and email sends (300-800ms) to evade automated spam filter detection.
                                     </p>
                                </div>

                                <div className="flex items-center gap-2 pt-1">
                                     <input
                                         id="schedule-toggle"
                                         type="checkbox"
                                         checked={isScheduling}
                                         onChange={(e) => setIsScheduling(e.target.checked)}
                                         disabled={isSending}
                                         className="w-4 h-4 rounded-none text-accent border-border"
                                     />
                                     <label htmlFor="schedule-toggle" className="text-xs font-medium text-ink cursor-pointer">
                                         Schedule campaign for a future date & time
                                     </label>
                                </div>

                                {isScheduling && (
                                    <div className="pt-1">
                                        <input
                                            type="datetime-local"
                                            value={scheduleDateTime}
                                            onChange={(e) => setScheduleDateTime(e.target.value)}
                                            min={getMinDateTime()}
                                            disabled={isSending}
                                            className="input py-1.5 text-xs"
                                        />
                                    </div>
                                )}

                                {scheduleError && (
                                    <p className="text-xs text-danger font-medium">{scheduleError}</p>
                                )}

                                <div className="flex gap-2 pt-2">
                                     <button onClick={onBack} disabled={isSending} className="btn btn-secondary text-xs px-4 py-2">
                                         Back
                                     </button>
                                     <button
                                         onClick={handleActionClick}
                                         disabled={isSending || (isScheduling && !scheduleDateTime) || recipientsForCampaign.length === 0}
                                         className="btn btn-primary flex-1 text-xs py-2.5 disabled:opacity-40 disabled:cursor-not-allowed shadow-xs"
                                     >
                                         {getButtonText()}
                                     </button>
                                 </div>
                             </div>
                         )}
                     </div>

                     {/* Right Column - Recipient Checklist & Progress */}
                     <div className="space-y-3">
                         <div className="flex items-center justify-between">
                             <h3 className="text-sm font-bold text-ink">Recipients Queue</h3>
                             {(isSending || isCampaignFinished) && (
                                 <span className="text-xs font-mono text-ink-muted">
                                     {sentCount + failedCount} / {campaignSize} ({progressPercentage}%)
                                 </span>
                             )}
                         </div>

                         {(isSending || isCampaignFinished) && (
                             <div className="h-2 rounded-xs overflow-hidden bg-neutral-200">
                                 <div
                                     className="h-full rounded-xs bg-accent transition-all duration-300"
                                     style={{ width: `${progressPercentage}%` }}
                                 />
                             </div>
                         )}

                         <div className="border border-border rounded-none overflow-hidden shadow-xs bg-card">
                            {!isSending && !isCampaignFinished && (
                                <div className="px-4 py-2.5 flex items-center justify-between border-b border-border bg-neutral-50">
                                    <div className="flex items-center gap-2">
                                        <input
                                            id="select-all"
                                            type="checkbox"
                                            checked={recipients.length > 0 && selectedEmails.length === recipients.length}
                                            onChange={handleSelectAll}
                                            className="w-4 h-4 rounded-none text-accent border-border"
                                        />
                                        <label htmlFor="select-all" className="text-xs font-semibold text-ink cursor-pointer">
                                            Select All Contacts
                                        </label>
                                    </div>
                                    <span className="text-xs font-mono text-ink-muted">
                                        {selectedEmails.length} of {recipients.length}
                                    </span>
                                </div>
                            )}

                            <div className="max-h-96 overflow-y-auto divide-y divide-border">
                                {displayList.map((recipient) => (
                                    <div key={recipient.email} className="px-4 py-3 flex items-center justify-between gap-3 hover:bg-neutral-50 transition-colors">
                                        <div className="flex items-center gap-3 min-w-0 flex-1">
                                            {!isSending && !isCampaignFinished && (
                                                <input
                                                    type="checkbox"
                                                    checked={selectedEmails.includes(recipient.email)}
                                                    onChange={() => handleRecipientSelect(recipient.email)}
                                                    className="w-4 h-4 rounded-none text-accent border-border flex-shrink-0"
                                                />
                                            )}
                                            <div className="min-w-0 flex-1">
                                                <p className="text-xs font-semibold text-ink truncate">
                                                    {recipient.fullName || recipient.firstName || 'Candidate'}
                                                </p>
                                                <p className="text-[11px] font-mono text-ink-muted truncate">
                                                    {recipient.email}
                                                </p>
                                            </div>
                                        </div>
                                        {(isSending || isCampaignFinished) && (
                                            <div className="flex-shrink-0">
                                                <StatusIndicator progress={sendProgress[recipient.email] || { status: EmailStatus.Queued }} />
                                            </div>
                                        )}
                                    </div>
                                ))}
                            </div>
                        </div>
                    </div>
                </div>
            ) : (
                <div className="text-center p-8 surface-card border border-red-200">
                    <p className="text-sm font-medium mb-4 text-danger">Please complete all previous steps first.</p>
                    <button onClick={onBack} className="btn btn-primary text-xs px-4 py-2">
                        Go Back
                    </button>
                </div>
            )}
        </div>
    );
};

export default ReviewAndSend;
