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
 * Displays email open and click tracking statistics
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
            <div className="glass-card p-6 animate-pulse">
                <div className="h-6 bg-white/10 rounded w-1/3 mb-4"></div>
                <div className="grid grid-cols-4 gap-4">
                    {[1, 2, 3, 4].map(i => (
                        <div key={i} className="h-16 bg-white/10 rounded"></div>
                    ))}
                </div>
            </div>
        );
    }

    if (error) {
        return (
            <div className="glass-card p-6 text-center text-red-400">
                <p>{error}</p>
                <button onClick={loadTrackingData} className="mt-2 text-blue-400 hover:underline">
                    Retry
                </button>
            </div>
        );
    }

    if (!data || data.summary.totalSent === 0) {
        return (
            <div className="glass-card p-6 text-center text-gray-400">
                <p>No tracking data available yet.</p>
                <p className="text-sm mt-1">Send some emails to see tracking statistics.</p>
            </div>
        );
    }

    const { summary, details } = data;

    return (
        <div className="glass-card p-6">
            {/* Header */}
            <div className="flex items-center justify-between mb-6">
                <h3 className="text-lg font-semibold text-white flex items-center gap-2">
                    <svg className="w-5 h-5 text-purple-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
                    </svg>
                    Email Tracking
                </h3>
                <button
                    onClick={() => setShowDetails(!showDetails)}
                    className="text-sm text-gray-400 hover:text-white transition-colors"
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
                    color="blue"
                />
                <StatCard
                    label="Opened"
                    value={summary.totalOpened}
                    icon={<EyeIcon />}
                    color="green"
                />
                <StatCard
                    label="Open Rate"
                    value={`${summary.openRate}%`}
                    icon={<ChartIcon />}
                    color="purple"
                />
                <StatCard
                    label="Clicks"
                    value={summary.totalClicks}
                    icon={<ClickIcon />}
                    color="orange"
                />
            </div>

            {/* Open Rate Progress Bar */}
            <div className="mb-6">
                <div className="flex justify-between text-sm mb-2">
                    <span className="text-gray-400">Open Rate</span>
                    <span className="text-white font-medium">{summary.openRate}%</span>
                </div>
                <div className="h-2 bg-white/10 rounded-full overflow-hidden">
                    <div
                        className="h-full bg-gradient-to-r from-purple-500 to-blue-500 transition-all duration-500"
                        style={{ width: `${summary.openRate}%` }}
                    />
                </div>
            </div>

            {/* Details Table */}
            {showDetails && details.length > 0 && (
                <div className="overflow-x-auto">
                    <table className="w-full text-sm">
                        <thead>
                            <tr className="text-left text-gray-400 border-b border-white/10">
                                <th className="pb-3">Recipient</th>
                                <th className="pb-3">Status</th>
                                <th className="pb-3">Opens</th>
                                <th className="pb-3">Clicks</th>
                                <th className="pb-3">First Opened</th>
                            </tr>
                        </thead>
                        <tbody>
                            {details.map((detail, index) => (
                                <tr key={index} className="border-b border-white/5 hover:bg-white/5">
                                    <td className="py-3 text-white truncate max-w-[200px]">
                                        {detail.recipientEmail}
                                    </td>
                                    <td className="py-3">
                                        {detail.opened ? (
                                            <span className="px-2 py-1 text-xs rounded-full bg-green-500/20 text-green-400">
                                                Opened
                                            </span>
                                        ) : (
                                            <span className="px-2 py-1 text-xs rounded-full bg-gray-500/20 text-gray-400">
                                                Not Opened
                                            </span>
                                        )}
                                    </td>
                                    <td className="py-3 text-white">{detail.openCount}</td>
                                    <td className="py-3 text-white">{detail.clicks}</td>
                                    <td className="py-3 text-gray-400">
                                        {detail.firstOpenedAt
                                            ? new Date(detail.firstOpenedAt).toLocaleString()
                                            : '-'}
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
const StatCard: React.FC<{
    label: string;
    value: string | number;
    icon: React.ReactNode;
    color: 'blue' | 'green' | 'purple' | 'orange';
}> = ({ label, value, icon, color }) => {
    const colors = {
        blue: 'from-blue-500 to-indigo-600',
        green: 'from-green-500 to-emerald-600',
        purple: 'from-purple-500 to-violet-600',
        orange: 'from-amber-500 to-orange-600',
    };

    return (
        <div className="bg-white/5 rounded-lg p-4">
            <div className="flex items-center gap-3">
                <div className={`p-2 rounded-lg bg-gradient-to-br ${colors[color]} text-white`}>
                    {icon}
                </div>
                <div>
                    <p className="text-xs text-gray-400 uppercase">{label}</p>
                    <p className="text-xl font-bold text-white">{value}</p>
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
