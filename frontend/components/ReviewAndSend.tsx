import React, { useState, useEffect } from 'react';
import { Credentials, Recipient, EmailTemplate, EmailStatus, SendProgressState } from '../types';

interface ReviewAndSendProps {
    credentials: Credentials | null;
    recipients: Recipient[];
    emailTemplate: EmailTemplate;
    sendProgress: Record<string, SendProgressState>;
    isSending: boolean;
    isCampaignFinished: boolean;
    scheduledTime: Date | null;
    onScheduleOrSend: (config: { time: Date | null, batchSize: number, batchDelay: number, recipientsToSend: Recipient[] }) => void;
    onCancelSchedule: () => void;
    onBack: () => void;
    onReset: () => void;
}

const StatusIndicator: React.FC<{ progress: SendProgressState }> = ({ progress }) => {
    if (!progress) return null;
    const { status, error } = progress;

    switch (status) {
        case EmailStatus.Queued:
            return <span className="flex items-center text-xs font-medium" style={{ color: '#64748b' }}><span className="w-2 h-2 rounded-full mr-2" style={{ background: '#64748b' }} />Queued</span>;
        case EmailStatus.Sending:
            return <span className="flex items-center text-xs font-medium" style={{ color: '#06b6d4' }}><span className="w-2 h-2 rounded-full mr-2 animate-pulse" style={{ background: '#06b6d4' }} />Sending...</span>;
        case EmailStatus.Sent:
            return <span className="flex items-center text-xs font-medium" style={{ color: '#14b8a6' }}><svg className="w-4 h-4 mr-1" fill="currentColor" viewBox="0 0 20 20"><path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" /></svg>Sent</span>;
        case EmailStatus.Failed:
            return <span title={error || 'Error'} className="flex items-center text-xs font-medium cursor-help" style={{ color: '#fb7185' }}><svg className="w-4 h-4 mr-1" fill="currentColor" viewBox="0 0 20 20"><path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" /></svg>Failed</span>;
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
    const progressPercentage = campaignSize > 0 ? ((sentCount + failedCount) / campaignSize) * 100 : 0;

    const getMinDateTime = () => {
        const now = new Date();
        now.setMinutes(now.getMinutes() + 1);
        return now.toISOString().slice(0, 16);
    };

    const handleActionClick = () => {
        const config = { time: null as Date | null, batchSize: Math.max(1, batchSize), batchDelay: Math.max(0, batchDelay), recipientsToSend: recipientsForCampaign };
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
    const getButtonText = () => isSending ? 'Sending...' : isScheduling ? 'Schedule Campaign' : `Send ${recipientsForCampaign.length} Emails`;

    return (
        <div className="max-w-4xl mx-auto fade-in">
            {/* Header */}
            <div className="text-center mb-8">
                <div className="inline-flex items-center justify-center w-14 h-14 rounded-2xl mb-4"
                    style={{ background: 'linear-gradient(135deg, rgba(6, 182, 212, 0.15), rgba(59, 130, 246, 0.15))', border: '1px solid rgba(6, 182, 212, 0.25)' }}>
                    <svg className="w-7 h-7" style={{ color: '#22d3ee' }} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                </div>
                <h2 className="text-2xl font-bold mb-2" style={{ color: '#f1f5f9' }}>Review & Send</h2>
                <p style={{ color: '#94a3b8', fontSize: '0.9375rem' }}>
                    {isCampaignFinished ? 'Campaign complete!' : scheduledTime ? 'Campaign scheduled.' : 'Final review before sending'}
                </p>
            </div>

            {isReady ? (
                <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    {/* Left Column - Summary */}
                    <div className="space-y-6">
                        <div className="rounded-xl p-5" style={{ background: 'rgba(148, 163, 184, 0.04)', border: '1px solid rgba(148, 163, 184, 0.12)' }}>
                            <h3 className="text-lg font-semibold mb-4 flex items-center gap-2" style={{ color: '#f1f5f9' }}>
                                <svg className="w-5 h-5" style={{ color: '#22d3ee' }} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2" />
                                </svg>
                                Campaign Summary
                            </h3>
                            <div className="space-y-3">
                                <div className="flex justify-between"><span style={{ color: '#94a3b8' }}>Sender</span><span style={{ color: '#f1f5f9' }}>{credentials?.email}</span></div>
                                <div className="flex justify-between"><span style={{ color: '#94a3b8' }}>Recipients</span><span style={{ color: '#f1f5f9' }}>{recipientsForCampaign.length} selected</span></div>
                                <div className="flex justify-between"><span style={{ color: '#94a3b8' }}>Subject</span><span className="truncate max-w-[200px]" style={{ color: '#f1f5f9' }}>{emailTemplate.subject}</span></div>
                                {emailTemplate.attachments && emailTemplate.attachments.length > 0 && (
                                    <div className="flex justify-between"><span style={{ color: '#94a3b8' }}>Attachments</span><span style={{ color: '#f1f5f9' }}>{emailTemplate.attachments.length} files</span></div>
                                )}
                            </div>
                            <div className="mt-4 pt-4" style={{ borderTop: '1px solid rgba(148, 163, 184, 0.1)' }}>
                                <p className="text-sm mb-2" style={{ color: '#64748b' }}>Preview</p>
                                <div className="p-3 rounded-lg max-h-32 overflow-y-auto text-sm" style={{ background: 'rgba(0,0,0,0.2)', color: '#94a3b8' }} dangerouslySetInnerHTML={{ __html: emailTemplate.body }} />
                            </div>
                        </div>

                        {isCampaignFinished ? (
                            <div className="rounded-xl p-5" style={{ background: 'rgba(20, 184, 166, 0.08)', border: '1px solid rgba(20, 184, 166, 0.2)' }}>
                                <h4 className="font-semibold mb-4" style={{ color: '#f1f5f9' }}>Campaign Results</h4>
                                <div className="grid grid-cols-2 gap-4">
                                    <div className="text-center p-4 rounded-lg" style={{ background: 'rgba(20, 184, 166, 0.15)' }}>
                                        <p className="text-3xl font-bold" style={{ color: '#14b8a6' }}>{sentCount}</p>
                                        <p className="text-sm" style={{ color: '#5eead4' }}>Sent</p>
                                    </div>
                                    <div className="text-center p-4 rounded-lg" style={{ background: 'rgba(244, 63, 94, 0.15)' }}>
                                        <p className="text-3xl font-bold" style={{ color: '#fb7185' }}>{failedCount}</p>
                                        <p className="text-sm" style={{ color: '#fda4af' }}>Failed</p>
                                    </div>
                                </div>
                                <button onClick={onReset} className="w-full mt-4 px-6 py-3 rounded-xl font-semibold text-white text-sm transition-all hover:scale-[1.02] active:scale-[0.98]" style={{ background: 'linear-gradient(135deg, #06b6d4, #3b82f6)', boxShadow: '0 4px 15px rgba(6, 182, 212, 0.35)' }}>
                                    Start New Campaign
                                </button>
                            </div>
                        ) : scheduledTime && !isSending ? (
                            <div className="rounded-xl p-5" style={{ background: 'rgba(6, 182, 212, 0.08)', border: '1px solid rgba(6, 182, 212, 0.2)' }}>
                                <p className="text-center mb-4" style={{ color: '#67e8f9' }}>Scheduled for:<br /><strong style={{ color: '#f1f5f9' }}>{scheduledTime.toLocaleString()}</strong></p>
                                <div className="flex gap-3">
                                    <button onClick={() => onScheduleOrSend({ time: null, batchSize, batchDelay, recipientsToSend: recipientsForCampaign })} className="flex-1 py-2 rounded-lg font-medium text-white text-sm" style={{ background: 'linear-gradient(135deg, #06b6d4, #3b82f6)' }}>Send Now</button>
                                    <button onClick={onCancelSchedule} className="flex-1 py-2 rounded-lg font-medium text-sm" style={{ background: 'rgba(148, 163, 184, 0.06)', border: '1px solid rgba(148, 163, 184, 0.15)', color: '#94a3b8' }}>Cancel</button>
                                </div>
                            </div>
                        ) : (
                            <div className="rounded-xl p-5" style={{ background: 'rgba(148, 163, 184, 0.04)', border: '1px solid rgba(148, 163, 184, 0.12)' }}>
                                <h4 className="font-semibold mb-4" style={{ color: '#f1f5f9' }}>Sending Options</h4>
                                <div className="grid grid-cols-2 gap-4 mb-4">
                                    <div>
                                        <label className="block text-sm mb-2" style={{ color: '#94a3b8', letterSpacing: '0.03em' }}>Batch Size</label>
                                        <input type="number" value={batchSize} onChange={(e) => setBatchSize(parseInt(e.target.value, 10) || 1)} min="1" disabled={isSending} />
                                    </div>
                                    <div>
                                        <label className="block text-sm mb-2" style={{ color: '#94a3b8', letterSpacing: '0.03em' }}>Delay (sec)</label>
                                        <input type="number" value={batchDelay} onChange={(e) => setBatchDelay(parseInt(e.target.value, 10) || 0)} min="0" disabled={isSending} />
                                    </div>
                                </div>
                                <div className="flex items-center gap-3 mb-4">
                                    <input id="schedule-toggle" type="checkbox" checked={isScheduling} onChange={(e) => setIsScheduling(e.target.checked)} disabled={isSending} className="w-4 h-4 rounded" />
                                    <label htmlFor="schedule-toggle" className="text-sm" style={{ color: '#94a3b8' }}>Schedule for later</label>
                                </div>
                                {isScheduling && (
                                    <input type="datetime-local" value={scheduleDateTime} onChange={(e) => setScheduleDateTime(e.target.value)} min={getMinDateTime()} disabled={isSending} className="w-full mb-4" />
                                )}
                                {scheduleError && <p className="text-sm mb-4" style={{ color: '#fb7185' }}>{scheduleError}</p>}
                                <div className="flex gap-3">
                                    <button onClick={onBack} disabled={isSending} className="px-6 py-3 rounded-xl font-medium text-sm transition-all" style={{ border: '1px solid rgba(148, 163, 184, 0.15)', color: '#94a3b8' }}>← Back</button>
                                    <button onClick={handleActionClick} disabled={isSending || (isScheduling && !scheduleDateTime) || recipientsForCampaign.length === 0} className="flex-1 py-3 rounded-xl font-semibold text-white text-sm transition-all disabled:opacity-40 hover:scale-[1.02] active:scale-[0.98]" style={{ background: 'linear-gradient(135deg, #06b6d4, #3b82f6)', boxShadow: '0 4px 20px rgba(6, 182, 212, 0.35)' }}>
                                        {getButtonText()}
                                    </button>
                                </div>
                            </div>
                        )}
                    </div>

                    {/* Right Column - Recipients */}
                    <div className="space-y-4">
                        <div className="flex items-center justify-between">
                            <h3 className="text-lg font-semibold" style={{ color: '#f1f5f9' }}>Recipients</h3>
                            {(isSending || isCampaignFinished) && (
                                <span className="text-sm" style={{ color: '#94a3b8' }}>{sentCount + failedCount} / {campaignSize}</span>
                            )}
                        </div>

                        {(isSending || isCampaignFinished) && (
                            <div className="h-1.5 rounded-full overflow-hidden" style={{ background: 'rgba(148, 163, 184, 0.1)' }}>
                                <div className="h-full transition-all duration-500 rounded-full" style={{ width: `${progressPercentage}%`, background: 'linear-gradient(90deg, #06b6d4, #3b82f6)' }} />
                            </div>
                        )}

                        <div className="rounded-xl overflow-hidden" style={{ background: 'rgba(148, 163, 184, 0.04)', border: '1px solid rgba(148, 163, 184, 0.12)' }}>
                            {!isSending && !isCampaignFinished && (
                                <div className="px-4 py-3 flex items-center justify-between" style={{ borderBottom: '1px solid rgba(148, 163, 184, 0.1)' }}>
                                    <div className="flex items-center gap-2">
                                        <input id="select-all" type="checkbox" checked={recipients.length > 0 && selectedEmails.length === recipients.length} onChange={handleSelectAll} className="w-4 h-4 rounded" />
                                        <label htmlFor="select-all" className="text-sm" style={{ color: '#94a3b8' }}>Select All</label>
                                    </div>
                                    <span className="text-sm" style={{ color: '#64748b' }}>{selectedEmails.length} / {recipients.length}</span>
                                </div>
                            )}
                            <div className="max-h-80 overflow-y-auto">
                                {displayList.map((recipient) => (
                                    <div key={recipient.email} className="px-4 py-3 flex items-center justify-between" style={{ borderBottom: '1px solid rgba(148, 163, 184, 0.06)' }}>
                                        <div className="flex items-center gap-3">
                                            {!isSending && !isCampaignFinished && (
                                                <input type="checkbox" checked={selectedEmails.includes(recipient.email)} onChange={() => handleRecipientSelect(recipient.email)} className="w-4 h-4 rounded" />
                                            )}
                                            <div>
                                                <p className="text-sm font-medium" style={{ color: '#f1f5f9' }}>{recipient.fullName}</p>
                                                <p className="text-xs" style={{ color: '#64748b' }}>{recipient.email}</p>
                                            </div>
                                        </div>
                                        {(isSending || isCampaignFinished) && <StatusIndicator progress={sendProgress[recipient.email] || { status: EmailStatus.Queued }} />}
                                    </div>
                                ))}
                            </div>
                        </div>
                    </div>
                </div>
            ) : (
                <div className="text-center p-8 rounded-xl" style={{ background: 'rgba(244, 63, 94, 0.08)', border: '1px solid rgba(244, 63, 94, 0.2)' }}>
                    <p className="mb-4" style={{ color: '#fb7185' }}>Please complete all previous steps first.</p>
                    <button onClick={onBack} className="px-6 py-3 rounded-xl font-semibold text-white text-sm" style={{ background: 'linear-gradient(135deg, #06b6d4, #3b82f6)' }}>Go Back</button>
                </div>
            )}
        </div>
    );
};

export default ReviewAndSend;
