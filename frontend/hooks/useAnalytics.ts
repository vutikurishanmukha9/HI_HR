import { useState, useEffect, useCallback, useRef } from 'react';
import apiClient from '../services/api';
import { TrackingStats, TrackingDetail, PaginationMeta, CampaignOption } from '../types';

interface UseAnalyticsOptions {
    autoRefreshInterval?: number; // In milliseconds, 0 to disable
}

function deriveCampaignOptions(details: TrackingDetail[]): CampaignOption[] {
    const map = new Map<string, { id: string; subject?: string; sentAt?: string; count: number }>();
    details.forEach(d => {
        if (!d.campaignId) return;
        const existing = map.get(d.campaignId);
        if (!existing) {
            map.set(d.campaignId, {
                id: d.campaignId,
                subject: d.subject,
                sentAt: d.sentAt,
                count: 1
            });
        } else {
            existing.count += 1;
            if (!existing.subject && d.subject) existing.subject = d.subject;
            if (d.sentAt && (!existing.sentAt || new Date(d.sentAt) > new Date(existing.sentAt))) {
                existing.sentAt = d.sentAt;
            }
        }
    });

    return Array.from(map.values()).map((c, i, arr) => {
        const dateObj = c.sentAt ? new Date(c.sentAt) : null;
        const dateStr = dateObj && !isNaN(dateObj.getTime())
            ? dateObj.toLocaleDateString('en-US', { month: 'short', day: 'numeric' })
            : '';
        let subj = (c.subject || '').trim() || `Campaign #${arr.length - i}`;
        if (subj.length > 34) subj = subj.substring(0, 34) + '...';
        const countSuffix = `${c.count} ${c.count === 1 ? 'email' : 'emails'}`;
        const label = dateStr ? `${subj} (${dateStr} · ${countSuffix})` : `${subj} (${countSuffix})`;
        return {
            id: c.id,
            label,
            subject: c.subject,
            sentAt: c.sentAt,
            totalCount: c.count
        };
    });
}

export function useAnalytics(options: UseAnalyticsOptions = { autoRefreshInterval: 60000 }) {
    const [stats, setStats] = useState<TrackingStats | null>(null);
    const [details, setDetails] = useState<TrackingDetail[]>([]);
    const [pagination, setPagination] = useState<PaginationMeta | null>(null);
    const [isLoading, setIsLoading] = useState<boolean>(true);
    const [error, setError] = useState<string | null>(null);

    // Filters and Pagination State
    const [page, setPage] = useState<number>(1);
    const [limit, setLimit] = useState<number>(10);
    const [selectedCampaign, setSelectedCampaign] = useState<string>('');

    // Available campaigns for filtering with human-readable labels
    const [availableCampaigns, setAvailableCampaigns] = useState<CampaignOption[]>([]);

    const fetchAnalytics = useCallback(async (silent = false) => {
        if (!silent) setIsLoading(true);
        setError(null);

        try {
            // We fetch details which also includes the summary stats for the current filter
            const response = await apiClient.getTrackingDetails(
                page,
                limit,
                selectedCampaign || undefined
            );

            if (response.success) {
                setStats(response.data.summary);
                setDetails(response.data.details);
                setPagination(response.data.pagination);

                if (response.data.campaigns && response.data.campaigns.length > 0) {
                    setAvailableCampaigns(response.data.campaigns);
                } else {
                    const derived = deriveCampaignOptions(response.data.details);
                    setAvailableCampaigns(prev => {
                        const merged = new Map<string, CampaignOption>();
                        prev.forEach(c => merged.set(c.id, c));
                        derived.forEach(c => merged.set(c.id, c));
                        return Array.from(merged.values());
                    });
                }
            } else {
                setError('Failed to fetch analytics data');
            }
        } catch (err: any) {
            setError(err.message || 'An error occurred while fetching analytics');
        } finally {
            if (!silent) setIsLoading(false);
        }
    }, [page, limit, selectedCampaign]);

    // Initial fetch and dependency trigger
    useEffect(() => {
        fetchAnalytics();
    }, [fetchAnalytics]);

    // Auto-refresh logic
    useEffect(() => {
        if (!options.autoRefreshInterval || options.autoRefreshInterval <= 0) return;

        const intervalId = setInterval(() => {
            fetchAnalytics(true); // Silent refresh
        }, options.autoRefreshInterval);

        return () => clearInterval(intervalId);
    }, [fetchAnalytics, options.autoRefreshInterval]);

    const handlePageChange = (newPage: number) => {
        if (pagination && newPage >= 1 && newPage <= pagination.totalPages) {
            setPage(newPage);
        }
    };

    const handleLimitChange = (newLimit: number) => {
        setLimit(newLimit);
        setPage(1); // Reset to first page
    };

    const handleCampaignChange = (campaignId: string) => {
        setSelectedCampaign(campaignId);
        setPage(1); // Reset to first page when filtering
    };

    const resetStats = async () => {
        setIsLoading(true);
        setError(null);
        try {
            const res = await apiClient.resetTrackingStats(selectedCampaign || undefined);
            if (res.success) {
                setStats({
                    totalSent: 0,
                    sentCount: 0,
                    failedCount: 0,
                    successRate: 0,
                });
                setDetails([]);
                setPagination(null);
                setAvailableCampaigns([]);
                await fetchAnalytics(false);
                return { success: true };
            }
            return { success: false, error: 'Failed to reset statistics' };
        } catch (err: any) {
            const msg = err.message || 'Error resetting statistics';
            setError(msg);
            return { success: false, error: msg };
        } finally {
            setIsLoading(false);
        }
    };

    return {
        stats,
        details,
        pagination,
        isLoading,
        error,
        page,
        limit,
        selectedCampaign,
        availableCampaigns,
        refresh: () => fetchAnalytics(false),
        silentRefresh: () => fetchAnalytics(true),
        resetStats,
        setPage: handlePageChange,
        setLimit: handleLimitChange,
        setSelectedCampaign: handleCampaignChange,
    };
}

export default useAnalytics;
