import React from 'react';

interface EmptyStateProps {
    type: 'recipients' | 'campaigns' | 'credentials' | 'general' | 'analytics';
    title?: string;
    message?: string;
    action?: {
        label: string;
        onClick: () => void;
    };
}

/**
 * EmptyState Component - Clean, AI-Free Semantic Badges
 * Uses crisp vector icons in light surface containers
 */
const EmptyState: React.FC<EmptyStateProps> = ({ type, title, message, action }) => {
    const icons: Record<string, React.ReactElement> = {
        recipients: (
            <div className="w-12 h-12 rounded-none bg-neutral-100 border border-border text-ink-muted flex items-center justify-center shadow-xs">
                <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.8}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z" />
                </svg>
            </div>
        ),
        campaigns: (
            <div className="w-12 h-12 rounded-none bg-accent-light border border-neutral-200 text-accent flex items-center justify-center shadow-xs">
                <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.8}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" />
                </svg>
            </div>
        ),
        credentials: (
            <div className="w-12 h-12 rounded-none bg-amber-50 border border-amber-200 text-amber-700 flex items-center justify-center shadow-xs">
                <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.8}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M15 7a2 2 0 012 2m4 0a6 6 0 01-7.743 5.743L11 17H9v2H7v2H4a1 1 0 01-1-1v-2.586a1 1 0 01.293-.707l5.964-5.964A6 6 0 1121 9z" />
                </svg>
            </div>
        ),
        general: (
            <div className="w-12 h-12 rounded-none bg-neutral-100 border border-border text-ink-muted flex items-center justify-center shadow-xs">
                <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.8}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M20 13V6a2 2 0 00-2-2H6a2 2 0 00-2 2v7m16 0v5a2 2 0 01-2 2H6a2 2 0 01-2-2v-5m16 0h-2.586a1 1 0 00-.707.293l-2.414 2.414a1 1 0 01-.707.293h-3.172a1 1 0 01-.707-.293l-2.414-2.414A1 1 0 006.586 13H4" />
                </svg>
            </div>
        ),
        analytics: (
            <div className="w-12 h-12 rounded-none bg-accent-light border border-neutral-200 text-accent flex items-center justify-center shadow-xs">
                <svg className="w-6 h-6" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={1.8}>
                    <path strokeLinecap="round" strokeLinejoin="round" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
                </svg>
            </div>
        ),
    };

    const defaults: Record<string, { title: string; message: string }> = {
        recipients: {
            title: 'No Recipients Configured',
            message: 'Upload an Excel or CSV spreadsheet with your prospective contact list to get started.',
        },
        campaigns: {
            title: 'No Campaigns Launched',
            message: 'Create your first email outreach campaign to contact talent prospects.',
        },
        credentials: {
            title: 'No Sender Connected',
            message: 'Connect your Gmail account with an App Password to dispatch emails.',
        },
        general: {
            title: 'No Records Found',
            message: 'There are no active records in this section.',
        },
        analytics: {
            title: 'No Tracking Data Yet',
            message: 'Dispatch your first cold email campaign to start receiving live delivery, open, and click analytics.',
        },
    };

    const { title: defaultTitle, message: defaultMessage } = defaults[type] || defaults.general;

    return (
        <div className="flex flex-col items-center justify-center py-10 px-4 text-center">
            <div className="mb-4">
                {icons[type] || icons.general}
            </div>
            <h3 className="text-base font-bold text-ink mb-1">
                {title || defaultTitle}
            </h3>
            <p className="text-xs text-ink-muted max-w-sm mb-5 leading-relaxed">
                {message || defaultMessage}
            </p>
            {action && (
                <button
                    onClick={action.onClick}
                    className="btn btn-primary text-xs px-4 py-2 shadow-xs"
                >
                    {action.label}
                </button>
            )}
        </div>
    );
};

export default EmptyState;
