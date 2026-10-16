import React from 'react';

interface DashboardStatsProps {
    totalSent: number;
    successRate: number;
    todaySent: number;
    pendingCount: number;
    onReset?: () => void;
}

/**
 * Dashboard Stats Component
 * Shows email campaign statistics with color-coded cards
 */
const DashboardStats: React.FC<DashboardStatsProps> = ({
    totalSent,
    successRate,
    todaySent,
    pendingCount,
    onReset,
}) => {
    const stats = [
        {
            label: 'Total Sent',
            value: totalSent,
            icon: (
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8" />
                </svg>
            ),
        },
        {
            label: 'Success Rate',
            value: `${successRate}%`,
            icon: (
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
            ),
        },
        {
            label: 'Today',
            value: todaySent,
            icon: (
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" />
                </svg>
            ),
        },
        {
            label: 'Pending',
            value: pendingCount,
            icon: (
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
            ),
        },
    ];

    return (
        <div className="mb-6 space-y-2">
            {onReset && (
                <div className="flex justify-between items-center px-1">
                    <span className="text-[11px] font-semibold uppercase tracking-wider text-ink-secondary">
                        Live Campaign Progress
                    </span>
                    <button
                        type="button"
                        onClick={onReset}
                        className="text-[11px] font-semibold text-danger hover:text-red-700 flex items-center gap-1 transition-colors"
                        title="Reset campaign and clear statistics"
                    >
                        <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                        </svg>
                        <span>Reset Campaign</span>
                    </button>
                </div>
            )}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
            {stats.map((stat) => (
                <div
                    key={stat.label}
                    className="surface-card p-4 border border-border shadow-xs"
                >
                    <div className="flex items-center gap-3">
                        <div className="p-2 rounded-none bg-accent-light text-accent border border-neutral-200 flex items-center justify-center flex-shrink-0">
                            {stat.icon}
                        </div>
                        <div className="min-w-0 flex-1">
                            <p className="text-[11px] uppercase tracking-wider font-semibold text-ink-secondary truncate">
                                {stat.label}
                            </p>
                            <p className="text-xl font-bold tabular-nums text-ink leading-tight truncate">
                                {stat.value}
                            </p>
                        </div>
                    </div>
                </div>
            ))}
            </div>
        </div>
    );
};

export default DashboardStats;
