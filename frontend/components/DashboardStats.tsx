import React from 'react';

interface DashboardStatsProps {
    totalSent: number;
    successRate: number;
    todaySent: number;
    pendingCount: number;
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
            gradient: 'linear-gradient(135deg, #6366f1, #818cf8)',
            bgColor: 'rgba(99, 102, 241, 0.08)',
            borderColor: 'rgba(99, 102, 241, 0.15)',
        },
        {
            label: 'Success Rate',
            value: `${successRate}%`,
            icon: (
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
            ),
            gradient: 'linear-gradient(135deg, #14b8a6, #10b981)',
            bgColor: 'rgba(20, 184, 166, 0.08)',
            borderColor: 'rgba(20, 184, 166, 0.15)',
        },
        {
            label: 'Today',
            value: todaySent,
            icon: (
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M8 7V3m8 4V3m-9 8h10M5 21h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z" />
                </svg>
            ),
            gradient: 'linear-gradient(135deg, #a855f7, #c084fc)',
            bgColor: 'rgba(168, 85, 247, 0.08)',
            borderColor: 'rgba(168, 85, 247, 0.15)',
        },
        {
            label: 'Pending',
            value: pendingCount,
            icon: (
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 8v4l3 3m6-3a9 9 0 11-18 0 9 9 0 0118 0z" />
                </svg>
            ),
            gradient: 'linear-gradient(135deg, #f59e0b, #fbbf24)',
            bgColor: 'rgba(245, 158, 11, 0.08)',
            borderColor: 'rgba(245, 158, 11, 0.15)',
        },
    ];

    return (
        <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mb-6">
            {stats.map((stat, index) => (
                <div
                    key={stat.label}
                    className="p-4 rounded-xl hover:scale-[1.02] transition-all duration-300"
                    style={{
                        background: stat.bgColor,
                        border: `1px solid ${stat.borderColor}`,
                        animationDelay: `${index * 100}ms`,
                    }}
                >
                    <div className="flex items-center gap-3">
                        <div
                            className="p-2 rounded-lg text-white"
                            style={{ background: stat.gradient }}
                        >
                            {stat.icon}
                        </div>
                        <div>
                            <p className="text-[10px] uppercase tracking-wider font-medium" style={{ color: '#64748b', letterSpacing: '0.08em' }}>{stat.label}</p>
                            <p className="text-xl font-bold" style={{ color: '#f1f5f9' }}>{stat.value}</p>
                        </div>
                    </div>
                </div>
            ))}
        </div>
    );
};

export default DashboardStats;
