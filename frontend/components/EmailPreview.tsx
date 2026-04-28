import React, { useState } from 'react';
import { Recipient, EmailTemplate } from '../types';

interface EmailPreviewProps {
    isOpen: boolean;
    onClose: () => void;
    emailTemplate: EmailTemplate;
    recipients: Recipient[];
    senderEmail: string;
}

/**
 * EmailPreview Component
 * Shows a preview of the email with placeholder values filled in
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

    // Replace placeholders with actual values
    const personalizeContent = (content: string): string => {
        return content
            .replace(/{fullName}/gi, selectedRecipient.fullName)
            .replace(/{companyName}/gi, selectedRecipient.companyName)
            .replace(/{jobTitle}/gi, selectedRecipient.jobTitle || '')
            .replace(/{email}/gi, selectedRecipient.email);
    };

    const personalizedSubject = personalizeContent(emailTemplate.subject);
    const personalizedBody = personalizeContent(emailTemplate.body);

    return (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ backgroundColor: 'rgba(0,0,0,0.75)' }}>
            <div className="glass-card w-full max-w-3xl max-h-[90vh] overflow-hidden flex flex-col">
                {/* Header */}
                <div className="flex items-center justify-between p-4 border-b border-white/10">
                    <h2 className="text-xl font-bold text-white">Email Preview</h2>
                    <button
                        onClick={onClose}
                        className="p-2 rounded-lg hover:bg-white/10 transition-colors"
                    >
                        <svg className="w-5 h-5 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                        </svg>
                    </button>
                </div>

                {/* Recipient Selector */}
                {recipients.length > 1 && (
                    <div className="p-4 bg-white/5 border-b border-white/10">
                        <label className="block text-sm text-gray-400 mb-2">
                            Preview for recipient:
                        </label>
                        <select
                            value={selectedRecipientIndex}
                            onChange={(e) => setSelectedRecipientIndex(Number(e.target.value))}
                            className="w-full px-3 py-2 rounded-lg bg-white/10 border border-white/20 text-white"
                        >
                            {recipients.map((r, i) => (
                                <option key={i} value={i} className="bg-gray-800">
                                    {r.fullName} ({r.email})
                                </option>
                            ))}
                        </select>
                    </div>
                )}

                {/* Email Preview */}
                <div className="flex-1 overflow-auto p-4">
                    {/* Email Header */}
                    <div className="mb-4 p-4 rounded-lg bg-white/5 space-y-2">
                        <div className="flex">
                            <span className="w-20 text-gray-400 text-sm">From:</span>
                            <span className="text-white">{senderEmail}</span>
                        </div>
                        <div className="flex">
                            <span className="w-20 text-gray-400 text-sm">To:</span>
                            <span className="text-white">{selectedRecipient.email}</span>
                        </div>
                        <div className="flex">
                            <span className="w-20 text-gray-400 text-sm">Subject:</span>
                            <span className="text-white font-medium">{personalizedSubject}</span>
                        </div>
                    </div>

                    {/* Email Body */}
                    <div className="p-4 rounded-lg bg-white border border-gray-200">
                        <div
                            className="prose prose-sm max-w-none text-gray-800"
                            dangerouslySetInnerHTML={{ __html: personalizedBody }}
                        />
                    </div>

                    {/* Attachments */}
                    {emailTemplate.attachments && emailTemplate.attachments.length > 0 && (
                        <div className="mt-4 p-4 rounded-lg bg-white/5">
                            <h4 className="text-sm font-medium text-gray-300 mb-2">
                                Attachments ({emailTemplate.attachments.length})
                            </h4>
                            <div className="space-y-1">
                                {emailTemplate.attachments.map((file, i) => (
                                    <div key={i} className="flex items-center gap-2 text-sm text-gray-400">
                                        <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15.172 7l-6.586 6.586a2 2 0 102.828 2.828l6.414-6.586a4 4 0 00-5.656-5.656l-6.415 6.585a6 6 0 108.486 8.486L20.5 13" />
                                        </svg>
                                        <span>{file.name}</span>
                                        <span className="text-gray-500">
                                            ({(file.size / 1024).toFixed(1)} KB)
                                        </span>
                                    </div>
                                ))}
                            </div>
                        </div>
                    )}
                </div>

                {/* Footer */}
                <div className="p-4 border-t border-white/10 flex justify-end">
                    <button
                        onClick={onClose}
                        className="px-6 py-2 rounded-lg bg-gradient-to-r from-purple-500 to-blue-500 text-white font-medium hover:opacity-90 transition-opacity"
                    >
                        Close Preview
                    </button>
                </div>
            </div>
        </div>
    );
};

export default EmailPreview;
