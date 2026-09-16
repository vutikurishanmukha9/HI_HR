import { useState, useEffect, useCallback, useRef } from 'react';
import apiClient from '../services/api';
import { TrackingStats, TrackingDetail, PaginationMeta } from '../types';

interface UseAnalyticsOptions {
    autoRefreshInterval?: number; // In milliseconds, 0 to disable
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

    // Available campaigns for filtering (derived from details or fetched separately)
    // For now, we'll derive it from the fetched data, though a dedicated endpoint would be better for scale
    const [availableCampaigns, setAvailableCampaigns] = useState<string[]>([]);

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

                // Extract unique campaigns for the filter dropdown if we don't have them
                // Note: In a real production app with thousands of campaigns, we'd want a separate endpoint for this
                const uniqueCampaigns = Array.from(new Set(
                    response.data.details
                        .map(d => d.campaignId)
                        .filter(Boolean) as string[]
                ));

                setAvailableCampaigns(prev => {
                    const combined = new Set([...prev, ...uniqueCampaigns]);
                    return Array.from(combined);
                });
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
