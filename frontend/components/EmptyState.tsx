import React from 'react';

interface EmptyStateProps {
    type: 'recipients' | 'campaigns' | 'credentials' | 'general';
    title?: string;
    message?: string;
    action?: {
        label: string;
        onClick: () => void;
    };
}

/**
 * Empty State Component
 * Shows illustrations for empty lists with call-to-action
 */
const EmptyState: React.FC<EmptyStateProps> = ({ type, title, message, action }) => {
    const illustrations: Record<string, React.ReactElement> = {
        recipients: (
            <svg className="w-32 h-32 text-gray-600" viewBox="0 0 200 200" fill="none">
                <circle cx="100" cy="100" r="80" stroke="currentColor" strokeWidth="2" strokeDasharray="8 4" />
                <circle cx="70" cy="80" r="20" fill="currentColor" opacity="0.3" />
                <circle cx="130" cy="80" r="20" fill="currentColor" opacity="0.3" />
                <circle cx="100" cy="120" r="25" fill="currentColor" opacity="0.3" />
                <path d="M60 150 Q100 180 140 150" stroke="currentColor" strokeWidth="3" strokeLinecap="round" fill="none" opacity="0.5" />
            </svg>
        ),
        campaigns: (
            <svg className="w-32 h-32 text-gray-600" viewBox="0 0 200 200" fill="none">
                <rect x="40" y="50" width="120" height="100" rx="8" stroke="currentColor" strokeWidth="2" />
                <path d="M40 70 L100 110 L160 70" stroke="currentColor" strokeWidth="2" fill="none" />
                <circle cx="160" cy="50" r="20" fill="currentColor" opacity="0.3" />
                <text x="153" y="56" fill="currentColor" fontSize="20" fontWeight="bold">+</text>
            </svg>
        ),
        credentials: (
            <svg className="w-32 h-32 text-gray-600" viewBox="0 0 200 200" fill="none">
                <rect x="50" y="60" width="100" height="80" rx="8" stroke="currentColor" strokeWidth="2" />
                <circle cx="100" cy="95" r="15" stroke="currentColor" strokeWidth="2" />
                <path d="M100 110 L100 125" stroke="currentColor" strokeWidth="2" />
                <rect x="90" y="120" width="20" height="15" rx="2" fill="currentColor" opacity="0.3" />
            </svg>
        ),
        general: (
            <svg className="w-32 h-32 text-gray-600" viewBox="0 0 200 200" fill="none">
                <circle cx="100" cy="100" r="60" stroke="currentColor" strokeWidth="2" strokeDasharray="8 4" />
                <path d="M80 90 L120 90 M80 110 L110 110" stroke="currentColor" strokeWidth="3" strokeLinecap="round" opacity="0.5" />
            </svg>
        ),
    };

    const defaults: Record<string, { title: string; message: string }> = {
        recipients: {
            title: 'No Recipients Yet',
            message: 'Upload an Excel file with your recipient list to get started.',
        },
        campaigns: {
            title: 'No Campaigns',
            message: 'Create your first email campaign to reach out to potential employers.',
        },
        credentials: {
            title: 'No Email Connected',
            message: 'Add your Gmail credentials to start sending emails.',
        },
        general: {
            title: 'Nothing Here',
            message: 'This section is empty.',
        },
    };

    const { title: defaultTitle, message: defaultMessage } = defaults[type] || defaults.general;

    return (
        <div className="flex flex-col items-center justify-center py-12 px-4 text-center">
            <div className="mb-6 opacity-50">
                {illustrations[type]}
            </div>
            <h3 className="text-xl font-semibold text-white mb-2">
                {title || defaultTitle}
            </h3>
            <p className="text-gray-400 max-w-sm mb-6">
                {message || defaultMessage}
            </p>
            {action && (
                <button
                    onClick={action.onClick}
                    className="px-6 py-3 bg-gradient-to-r from-purple-500 to-blue-500 text-white font-medium rounded-lg hover:opacity-90 transition-opacity"
                >
                    {action.label}
                </button>
            )}
        </div>
    );
};

export default EmptyState;
