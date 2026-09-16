import React, { useState } from 'react';
import DOMPurify from 'dompurify';
import { Recipient, EmailTemplate } from '../types';

interface EmailPreviewProps {
    isOpen: boolean;
    onClose: () => void;
    emailTemplate: EmailTemplate;
    recipients: Recipient[];
    senderEmail: string;
}

/**
 * EmailPreview Component - Light Theme Modal
 * Clean simulation of actual delivered email with filled variables
 */
const EmailPreview: React.FC<EmailPreviewProps> = ({
    isOpen,
    onClose,
    emailTemplate,
    recipients,
    senderEmail,
}) => {
    const [selectedRecipientIndex, setSelectedRecipientIndex] = useState(0);

    if (!isOpen) return null;

    const selectedRecipient = recipients[selectedRecipientIndex] || {
        email: 'recipient@example.com',
        fullName: 'John Doe',
        companyName: 'Example Corp',
        jobTitle: 'HR Manager',
    };

    // Replace placeholders with actual values dynamically
    const personalizeContent = (content: string): string => {
        let personalized = content;
        Object.entries(selectedRecipient).forEach(([key, value]) => {
            const regex = new RegExp(`{${key}}`, 'gi');
            personalized = personalized.replace(regex, String(value || ''));
        });
        return personalized;
    };

    const personalizedSubject = personalizeContent(emailTemplate.subject);
    const personalizedBody = personalizeContent(emailTemplate.body);

    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/40 backdrop-blur-xs animate-fadeIn">
            <div className="surface-card w-full max-w-2xl max-h-[85vh] overflow-hidden flex flex-col border border-border shadow-xl bg-card rounded-2xl">
                {/* Header */}
                <div className="flex items-center justify-between px-5 py-4 border-b border-border">
                    <div className="flex items-center gap-2">
                        <div className="w-7 h-7 rounded-lg bg-accent-light text-accent flex items-center justify-center">
                            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
                            </svg>
                        </div>
                        <h2 className="text-base font-bold text-ink">Email Preview</h2>
                    </div>
                    <button
                        onClick={onClose}
                        className="p-1.5 rounded-lg text-ink-muted hover:text-ink hover:bg-neutral-100 transition-colors"
                    >
                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                        </svg>
                    </button>
                </div>

                {/* Recipient Selector */}
                {recipients.length > 1 && (
                    <div className="px-5 py-3 border-b border-border bg-neutral-50">
                        <label className="block text-[11px] font-semibold uppercase tracking-wider text-ink-secondary mb-1.5">
                            Simulate for recipient:
                        </label>
                        <div className="relative">
                            <select
                                value={selectedRecipientIndex}
                                onChange={(e) => setSelectedRecipientIndex(Number(e.target.value))}
                                className="input py-1.5 text-xs bg-white appearance-none cursor-pointer input-icon-right !pr-8"
                                style={{ paddingRight: '2rem' }}
                            >
                                {recipients.map((r, i) => {
                                    const displayName = r.fullName || r.firstName || r.name || r.email;
                                    return (
                                        <option key={i} value={i}>
                                            {displayName !== r.email ? `${displayName} (${r.email})` : r.email}
                                        </option>
                                    );
                                })}
                            </select>
                            <div className="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-ink-muted">
                                <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" /></svg>
                            </div>
                        </div>
                    </div>
                )}

                {/* Email Content Preview */}
                <div className="flex-1 overflow-auto p-5 space-y-4">
                    {/* Header Metadata */}
                    <div className="p-3.5 rounded-xl bg-neutral-50 border border-border text-xs space-y-1.5">
                        <div className="flex items-center gap-2">
                            <span className="w-16 flex-shrink-0 font-semibold text-ink-muted">From:</span>
                            <span className="font-mono text-ink truncate">{senderEmail}</span>
                        </div>
                        <div className="flex items-center gap-2">
                            <span className="w-16 flex-shrink-0 font-semibold text-ink-muted">To:</span>
                            <span className="font-mono text-ink truncate">{selectedRecipient.email}</span>
                        </div>
                        <div className="flex items-start gap-2">
                            <span className="w-16 flex-shrink-0 font-semibold text-ink-muted mt-0.5">Subject:</span>
                            <span className="font-semibold text-ink break-words min-w-0 flex-1">{personalizedSubject || '(Empty Subject)'}</span>
                        </div>
                    </div>

                    {/* Body */}
                    <div className="p-5 rounded-xl border border-border bg-white min-h-[160px]">
                        <div
                            className="text-sm text-ink leading-relaxed prose prose-sm max-w-none"
                            dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(personalizedBody || '<p class="text-ink-muted italic">No content composed.</p>') }}
                        />
                    </div>

                    {/* Attachments */}
                    {emailTemplate.attachments && emailTemplate.attachments.length > 0 && (
                        <div className="p-3.5 rounded-xl bg-neutral-50 border border-border">
                            <h4 className="text-xs font-semibold text-ink-secondary uppercase tracking-wider mb-2">
                                Attachments ({emailTemplate.attachments.length})
                            </h4>
                            <div className="space-y-1.5">
                                {emailTemplate.attachments.map((file, i) => (
                                    <div key={i} className="flex items-center justify-between text-xs p-2 rounded-lg bg-white border border-border">
                                        <div className="flex items-center gap-2">
                                            <svg className="w-4 h-4 text-accent" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15.172 7l-6.586 6.586a2 2 0 102.828 2.828l6.414-6.586a4 4 0 00-5.656-5.656l-6.415 6.585a6 6 0 108.486 8.486L20.5 13" />
                                            </svg>
                                             <span className="font-medium text-ink">{file.name}</span>
                                        </div>
                                        <span className="text-ink-muted text-[11px]">
                                            {(file.size / 1024).toFixed(1)} KB
                                        </span>
                                    </div>
                                ))}
                            </div>
                        </div>
                    )}
                </div>

                {/* Footer */}
                <div className="px-5 py-3.5 border-t border-border bg-neutral-50 flex justify-end">
                    <button
                        onClick={onClose}
                        className="btn btn-primary text-xs px-4 py-2"
                    >
                        Done Previewing
                    </button>
                </div>
            </div>
        </div>
    );
};

export default EmailPreview;
