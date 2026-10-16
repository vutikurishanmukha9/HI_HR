import React from 'react';
import { TrackingDetail, PaginationMeta } from '../../types';

interface TrackingTableProps {
    details: TrackingDetail[];
    pagination: PaginationMeta | null;
    isLoading: boolean;
    onPageChange: (page: number) => void;
}

/**
 * TrackingTable Component
 * Renders campaign recipient tracking detail rows.
 */
const TrackingTable: React.FC<TrackingTableProps> = ({ details, pagination, isLoading, onPageChange }) => {
    
    const formatDate = (dateString?: string) => {
        if (!dateString) return '-';
        const date = new Date(dateString);
        return new Intl.DateTimeFormat('en-US', {
            month: 'short',
            day: 'numeric',
            hour: 'numeric',
            minute: '2-digit'
        }).format(date);
    };

    if (isLoading) {
        return (
            <div className="surface-card p-8 flex justify-center items-center h-64">
                <div className="spinner border-t-accent"></div>
            </div>
        );
    }

    return (
        <div className="surface-card overflow-hidden border border-border shadow-xs">
            <div className="overflow-x-auto">
                <table className="table w-full">
                    <thead>
                        <tr>
                            <th className="text-[11px] whitespace-nowrap">Recipient</th>
                            <th className="text-[11px] whitespace-nowrap">Subject</th>
                            <th className="text-[11px] whitespace-nowrap">Delivery Status</th>
                            <th className="text-[11px] whitespace-nowrap">Dispatched At</th>
                            <th className="text-[11px] whitespace-nowrap">Notes / Diagnostics</th>
                        </tr>
                    </thead>
                    <tbody>
                        {details.map((row, i) => (
                            <tr key={i} className="hover:bg-neutral-50 transition-colors">
                                <td className="font-mono text-xs text-ink font-medium truncate max-w-[200px]" title={row.recipientEmail}>
                                    {row.recipientEmail}
                                </td>
                                <td>
                                    <div className="text-xs text-ink truncate max-w-[220px]" title={row.subject || row.campaignId}>
                                        {row.subject || 'Cold Outreach Email'}
                                    </div>
                                    {row.campaignId && (
                                        <div className="text-[10px] text-ink-muted font-mono truncate max-w-[180px]">
                                            ID: {row.campaignId.substring(0, 16)}...
                                        </div>
                                    )}
                                </td>
                                <td className="whitespace-nowrap">
                                    {row.status === 'sent' ? (
                                        <span className="badge badge-success rounded-none flex items-center gap-1 w-fit">
                                            <svg className="w-3 h-3 text-emerald-600 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M5 13l4 4L19 7" /></svg>
                                            <span>Delivered</span>
                                        </span>
                                    ) : (
                                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-none text-[11px] font-medium bg-red-50 text-red-700 border border-red-200 w-fit">
                                            <svg className="w-3 h-3 text-red-500 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
                                            <span>Failed</span>
                                        </span>
                                    )}
                                </td>
                                <td className="whitespace-nowrap text-xs text-ink-muted font-mono">
                                    {formatDate(row.sentAt)}
                                </td>
                                <td className="text-xs">
                                    {row.errorMessage ? (
                                        <span className="text-danger font-medium truncate max-w-[260px] block" title={row.errorMessage}>
                                            {row.errorMessage}
                                        </span>
                                    ) : (
                                        <span className="text-ink-muted">
                                            Sent cleanly via Gmail SMTP
                                        </span>
                                    )}
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>

            {/* Pagination Controls */}
            {pagination && pagination.totalPages > 1 && (
                <div className="px-5 py-3 flex items-center justify-between border-t border-border bg-neutral-50">
                    <div className="text-xs text-ink-muted">
                        Showing <span className="font-semibold text-ink">{((pagination.page - 1) * pagination.limit) + 1}</span> to <span className="font-semibold text-ink">{Math.min(pagination.page * pagination.limit, pagination.total)}</span> of <span className="font-semibold text-ink">{pagination.total}</span> recipients
                    </div>
                    <div className="flex gap-2">
                        <button
                            onClick={() => onPageChange(pagination.page - 1)}
                            disabled={pagination.page <= 1}
                            className="btn btn-secondary text-xs px-3 py-1.5 disabled:opacity-40"
                        >
                            Previous
                        </button>
                        <button
                            onClick={() => onPageChange(pagination.page + 1)}
                            disabled={pagination.page >= pagination.totalPages}
                            className="btn btn-secondary text-xs px-3 py-1.5 disabled:opacity-40"
                        >
                            Next
                        </button>
                    </div>
                </div>
            )}
        </div>
    );
};

export default TrackingTable;
