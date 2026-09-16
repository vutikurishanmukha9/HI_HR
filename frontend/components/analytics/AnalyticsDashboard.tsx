import React from 'react';
import useAnalytics from '../../hooks/useAnalytics';
import StatCard from './StatCard';
import TrackingTable from './TrackingTable';
import EmptyState from '../EmptyState';

/**
 * AnalyticsDashboard Component
 * Main page for campaign live tracking analytics under the Hallmark design system.
 */
const AnalyticsDashboard: React.FC = () => {
    const { 
        stats, 
        details, 
        pagination, 
        isLoading, 
        error, 
        selectedCampaign, 
        availableCampaigns,
        refresh, 
        resetStats,
        setPage, 
        setSelectedCampaign 
    } = useAnalytics({ autoRefreshInterval: 60000 }); // 60s auto refresh

    const [showResetConfirm, setShowResetConfirm] = React.useState(false);
    const [isResetting, setIsResetting] = React.useState(false);
    const [resetNotice, setResetNotice] = React.useState<string | null>(null);

    const handleConfirmReset = async () => {
        setIsResetting(true);
        const result = await resetStats();
        setIsResetting(false);
        setShowResetConfirm(false);
        if (result.success) {
            setResetNotice('All campaign tracking statistics have been cleared.');
            setTimeout(() => setResetNotice(null), 4000);
        }
    };

    const hasData = details.length > 0 || (stats && stats.totalSent > 0);

    return (
        <div className="space-y-6">
            {/* Header Area with Filters & Actions */}
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 surface-card p-4 sm:p-5">
                <div className="flex items-center gap-3">
                    <div className="w-10 h-10 rounded-xl bg-accent-light border border-neutral-200 text-accent flex items-center justify-center shadow-xs">
                        <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
                        </svg>
                    </div>
                    <div>
                        <h2 className="text-base font-bold text-ink">Outreach Delivery & Logs</h2>
                        <p className="text-xs text-ink-muted">Real-time email dispatch and delivery status</p>
                    </div>
                </div>

                <div className="flex items-center gap-2.5 w-full sm:w-auto">
                    {/* Campaign Filter */}
                    <div className="relative flex-1 sm:flex-none sm:w-64">
                        <select
                            value={selectedCampaign}
                            onChange={(e) => setSelectedCampaign(e.target.value)}
                            className="input py-2 text-xs bg-white input-icon-right !pr-9 appearance-none cursor-pointer"
                            style={{ paddingRight: '2.25rem' }}
                        >
                            <option value="">All Outreach Campaigns</option>
                            {availableCampaigns.map(camp => (
                                <option key={camp} value={camp}>
                                    {camp.substring(0, 30)}{camp.length > 30 ? '...' : ''}
                                </option>
                            ))}
                        </select>
                        <div className="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none text-ink-muted">
                            <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 9l-7 7-7-7" /></svg>
                        </div>
                    </div>

                    {/* Reset Statistics Button */}
                    <button 
                        onClick={() => setShowResetConfirm(true)}
                        disabled={isLoading || isResetting}
                        className="btn btn-secondary px-3 py-2 text-xs font-semibold text-danger hover:bg-red-50 hover:border-red-200 flex items-center gap-1.5 transition-colors disabled:opacity-50"
                        title="Clear all statistics and logs"
                    >
                        <svg className="w-4 h-4 text-danger" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                        </svg>
                        <span className="hidden sm:inline">Reset Stats</span>
                    </button>

                    {/* Refresh Button */}
                    <button 
                        onClick={refresh}
                        disabled={isLoading || isResetting}
                        className="btn btn-secondary p-2 text-ink-secondary hover:text-ink disabled:opacity-50"
                        title="Refresh Analytics"
                    >
                        <svg className={`w-4 h-4 ${isLoading ? 'animate-spin' : ''}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                        </svg>
                    </button>
                </div>
            </div>

            {/* Reset Confirmation Modal */}
            {showResetConfirm && (
                <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-neutral-900/40 backdrop-blur-xs animate-fadeIn">
                    <div className="surface-card p-6 rounded-2xl max-w-md w-full shadow-lg border border-border space-y-4 animate-scaleUp">
                        <div className="flex items-start gap-3">
                            <div className="w-10 h-10 rounded-xl bg-danger-light border border-red-200 text-danger flex items-center justify-center flex-shrink-0">
                                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                                </svg>
                            </div>
                            <div className="min-w-0 flex-1">
                                <h3 className="text-base font-bold text-ink">Reset All Statistics?</h3>
                                <p className="text-xs text-ink-muted mt-1 leading-relaxed">
                                    This will permanently clear all delivery metrics, open tracking records, click logs, and recipient activity{selectedCampaign ? ` for this campaign` : ''}.
                                </p>
                            </div>
                        </div>

                        <div className="flex justify-end gap-2 pt-2">
                            <button
                                type="button"
                                onClick={() => setShowResetConfirm(false)}
                                disabled={isResetting}
                                className="btn btn-secondary text-xs px-4 py-2"
                            >
                                Cancel
                            </button>
                            <button
                                type="button"
                                onClick={handleConfirmReset}
                                disabled={isResetting}
                                className="btn text-xs px-4 py-2 bg-red-600 hover:bg-red-700 text-white font-semibold rounded-lg shadow-xs flex items-center gap-1.5"
                            >
                                {isResetting ? (
                                    <>
                                        <div className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />
                                        <span>Resetting…</span>
                                    </>
                                ) : (
                                    <span>Yes, Reset Statistics</span>
                                )}
                            </button>
                        </div>
                    </div>
                </div>
            )}

            {resetNotice && (
                <div className="p-3.5 rounded-xl bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs font-medium animate-fadeIn flex items-center gap-2">
                    <svg className="w-4 h-4 text-emerald-600 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                        <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zm3.707-9.293a1 1 0 00-1.414-1.414L9 10.586 7.707 9.293a1 1 0 00-1.414 1.414l2 2a1 1 0 001.414 0l4-4z" clipRule="evenodd" />
                    </svg>
                    <span>{resetNotice}</span>
                </div>
            )}

            {error && (
                <div className="p-3.5 rounded-xl bg-danger-light border border-red-200 text-danger text-xs font-medium animate-fadeIn">
                    {error}
                </div>
            )}

            {!hasData && !isLoading ? (
                <div className="surface-card p-6 border border-border shadow-xs">
                    <EmptyState type="analytics" />
                </div>
            ) : (
                <>
                    {/* Stats Overview */}
                    {stats && (
                        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                            <StatCard 
                                title="Total Outreach"
                                value={stats.totalSent}
                                subtitle="Total recipients queued"
                                icon={<svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 19l9 2-9-18-9 18 9-2zm0 0v-8"></path></svg>}
                            />
                            <StatCard 
                                title="Delivered"
                                value={stats.sentCount}
                                subtitle={`${stats.successRate}% Success`}
                                icon={<svg className="w-5 h-5 text-emerald-600" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7"></path></svg>}
                            />
                            <StatCard 
                                title="Failed"
                                value={stats.failedCount}
                                subtitle={stats.failedCount > 0 ? "Requires review" : "Zero errors"}
                                icon={<svg className="w-5 h-5 text-red-500" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>}
                            />
                            <StatCard 
                                title="Delivery Rate"
                                value={`${stats.successRate}%`}
                                subtitle="Dispatched successfully"
                                icon={<svg className="w-5 h-5 text-accent" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>}
                            />
                        </div>
                    )}

                    {/* Tracking Details Table */}
                    <TrackingTable 
                        details={details} 
                        pagination={pagination}
                        isLoading={isLoading && details.length === 0} 
                        onPageChange={setPage}
                    />
                </>
            )}
        </div>
    );
};

export default AnalyticsDashboard;
