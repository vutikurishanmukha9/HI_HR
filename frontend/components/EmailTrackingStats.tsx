import React, { useEffect, useState } from 'react';
import api from '../services/api';

interface TrackingStats {
    totalSent: number;
    totalOpened: number;
    openRate: number;
    totalClicks: number;
    uniqueClicks: number;
}

interface TrackingDetail {
    recipientEmail: string;
    subject?: string;
    opened: boolean;
    openCount: number;
    firstOpenedAt?: string;
    lastOpenedAt?: string;
    clicks: number;
    links: Array<{
        url: string;
        clicks: number;
        firstClickedAt?: string;
    }>;
}

interface TrackingData {
    summary: TrackingStats;
    details: TrackingDetail[];
}

/**
 * EmailTrackingStats Component
 * Displays email open and click tracking statistics using the Hallmark design system.
 */
const EmailTrackingStats: React.FC<{ campaignId?: string }> = ({ campaignId }) => {
    const [data, setData] = useState<TrackingData | null>(null);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState<string | null>(null);
    const [showDetails, setShowDetails] = useState(false);

    useEffect(() => {
        loadTrackingData();
    }, [campaignId]);

    const loadTrackingData = async () => {
        try {
            setLoading(true);
            const params = campaignId ? `?campaignId=${campaignId}` : '';
            const response = await api.get<{ success: boolean; data: TrackingData }>(
                `/track/details${params}`
            );

            if (response.success && response.data) {
                setData(response.data);
            } else {
                setData(null);
            }
        } catch (err: any) {
            setError(err.message || 'Failed to load tracking data');
        } finally {
            setLoading(false);
        }
    };

    if (loading) {
        return (
            <div className="surface-card p-6 animate-shimmer rounded-none">
                <div className="h-5 bg-neutral-200 rounded-none w-1/3 mb-4"></div>
                <div className="grid grid-cols-4 gap-4">
                    {[1, 2, 3, 4].map(i => (
                        <div key={i} className="h-16 bg-neutral-100 rounded-none"></div>
                    ))}
                </div>
            </div>
        );
    }

    if (error) {
        return (
            <div className="surface-card p-6 text-center text-danger bg-danger-light border border-red-200">
                <p className="text-xs font-medium">{error}</p>
                <button onClick={loadTrackingData} className="mt-2 text-accent text-xs font-semibold hover:underline">
                    Retry
                </button>
            </div>
        );
    }

    if (!data || data.summary.totalSent === 0) {
        return (
            <div className="surface-card p-6 text-center text-ink-muted">
                <p className="text-sm font-medium">No tracking data available yet.</p>
                <p className="text-xs mt-1">Send some emails to see tracking statistics.</p>
            </div>
        );
    }

    const { summary, details } = data;

    return (
        <div className="surface-card p-6 border border-border shadow-xs">
            {/* Header */}
            <div className="flex items-center justify-between mb-6">
                <h3 className="text-base font-bold text-ink flex items-center gap-2">
                    <svg className="w-5 h-5 text-accent" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
                    </svg>
                    Email Delivery & Engagement
                </h3>
                <button
                    onClick={() => setShowDetails(!showDetails)}
                    className="btn btn-secondary text-xs px-3 py-1.5"
                >
                    {showDetails ? 'Hide Details' : 'Show Details'}
                </button>
            </div>

            {/* Stats Cards */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
                <StatCard
                    label="Sent"
                    value={summary.totalSent}
                    icon={<SendIcon />}
                />
                <StatCard
                    label="Opened"
                    value={summary.totalOpened}
                    icon={<EyeIcon />}
                />
                <StatCard
                    label="Open Rate"
                    value={`${summary.openRate}%`}
                    icon={<ChartIcon />}
                />
                <StatCard
                    label="Clicks"
                    value={summary.totalClicks}
                    icon={<ClickIcon />}
                />
            </div>

            {/* Open Rate Progress Bar */}
            <div className="mb-6 p-4 rounded-none bg-neutral-50 border border-border">
                <div className="flex justify-between text-xs mb-2">
                    <span className="font-semibold text-ink-secondary uppercase tracking-wider">Campaign Open Rate</span>
                    <span className="text-ink font-bold tabular-nums">{summary.openRate}%</span>
                </div>
                <div className="h-2 bg-neutral-200 rounded-none overflow-hidden">
                    <div
                        className="h-full bg-accent transition-all duration-500 rounded-none"
                        style={{ width: `${summary.openRate}%` }}
                    />
                </div>
            </div>

            {/* Details Table */}
            {showDetails && details.length > 0 && (
                <div className="overflow-x-auto border border-border rounded-none">
                    <table className="table w-full">
                        <thead>
                            <tr>
                                <th className="text-[11px] whitespace-nowrap">Recipient</th>
                                <th className="text-[11px] whitespace-nowrap">Status</th>
                                <th className="text-[11px] whitespace-nowrap">Opens</th>
                                <th className="text-[11px] whitespace-nowrap">Clicks</th>
                                <th className="text-[11px] whitespace-nowrap">First Opened</th>
                            </tr>
                        </thead>
                        <tbody>
                            {details.map((detail, index) => (
                                <tr key={index} className="hover:bg-neutral-50 transition-colors">
                                    <td className="font-mono text-xs text-ink truncate max-w-[200px]" title={detail.recipientEmail}>
                                        {detail.recipientEmail}
                                    </td>
                                    <td>
                                        {detail.opened ? (
                                            <span className="badge badge-success rounded-none">
                                                Opened
                                            </span>
                                        ) : (
                                            <span className="inline-flex items-center px-2 py-0.5 rounded-none text-[11px] font-medium bg-neutral-100 text-neutral-500 border border-neutral-200">
                                                Not Opened
                                            </span>
                                        )}
                                    </td>
                                    <td className="text-ink font-mono text-xs">{detail.openCount}</td>
                                    <td className="text-ink font-mono text-xs">{detail.clicks}</td>
                                    <td className="text-ink-muted font-mono text-xs whitespace-nowrap">
                                        {detail.firstOpenedAt
                                            ? new Date(detail.firstOpenedAt).toLocaleString()
                                            : '—'}
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            )}
        </div>
    );
};

// Helper Components
interface HelperStatCardProps {
    label: string;
    value: string | number;
    icon: React.ReactNode;
}

const StatCard: React.FC<HelperStatCardProps> = ({ label, value, icon }) => {
    return (
        <div className="bg-card border border-border rounded-none p-4 shadow-xs">
            <div className="flex items-center gap-3">
                <div className="p-2 rounded-none bg-accent-light text-accent flex items-center justify-center flex-shrink-0">
                    {icon}
                </div>
                <div className="min-w-0 flex-1">
                    <p className="text-[11px] font-semibold text-ink-secondary uppercase tracking-wider truncate">{label}</p>
                    <p className="text-xl font-bold text-ink tabular-nums leading-tight truncate">{value}</p>
                </div>
            </div>
        </div>
    );
};

// Icons
const SendIcon = () => (
    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8" />
    </svg>
);

const EyeIcon = () => (
    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
    </svg>
);

const ChartIcon = () => (
    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
    </svg>
);

const ClickIcon = () => (
    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 15l-2 5L9 9l11 4-5 2zm0 0l5 5M7.188 2.239l.777 2.897M5.136 7.965l-2.898-.777M13.95 4.05l-2.122 2.122m-5.657 5.656l-2.12 2.122" />
    </svg>
);

export default EmailTrackingStats;
