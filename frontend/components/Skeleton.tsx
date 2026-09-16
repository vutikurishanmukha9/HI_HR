import React from 'react';

/**
 * Skeleton Loader Components
 * Shimmer effect placeholders for loading states
 * Uses animate-shimmer class (subtle opacity pulse on paper-3 surface)
 */

// Base skeleton with shimmer
export const Skeleton: React.FC<{ className?: string }> = ({ className = '' }) => (
    <div
        className={`animate-shimmer rounded ${className}`}
    />
);

// Text skeleton
export const SkeletonText: React.FC<{ lines?: number; className?: string }> = ({ lines = 1, className = '' }) => (
    <div className={`space-y-2 ${className}`}>
        {Array.from({ length: lines }).map((_, i) => (
            <Skeleton key={i} className={`h-4 ${i === lines - 1 ? 'w-3/4' : 'w-full'}`} />
        ))}
    </div>
);

// Card skeleton
export const SkeletonCard: React.FC<{ className?: string }> = ({ className = '' }) => (
    <div className={`surface-card p-6 space-y-4 ${className}`}>
        <div className="flex items-center gap-4">
            <Skeleton className="w-12 h-12 rounded-xl" />
            <div className="flex-1 space-y-2">
                <Skeleton className="h-5 w-1/3" />
                <Skeleton className="h-4 w-1/2" />
            </div>
        </div>
        <SkeletonText lines={3} />
    </div>
);

// Table row skeleton
export const SkeletonTableRow: React.FC<{ columns?: number }> = ({ columns = 4 }) => (
    <tr>
        {Array.from({ length: columns }).map((_, i) => (
            <td key={i} className="p-4">
                <Skeleton className="h-4 w-full" />
            </td>
        ))}
    </tr>
);

// Stats card skeleton
export const SkeletonStats: React.FC = () => (
    <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
        {[1, 2, 3, 4].map((i) => (
            <div key={i} className="surface-card p-4 space-y-2">
                <Skeleton className="h-4 w-20" />
                <Skeleton className="h-8 w-16" />
            </div>
        ))}
    </div>
);

// Avatar skeleton
export const SkeletonAvatar: React.FC<{ size?: 'sm' | 'md' | 'lg' }> = ({ size = 'md' }) => {
    const sizes = {
        sm: 'w-8 h-8',
        md: 'w-10 h-10',
        lg: 'w-12 h-12',
    };
    return <Skeleton className={`${sizes[size]} rounded-full`} />;
};

// Button skeleton
export const SkeletonButton: React.FC<{ width?: string }> = ({ width = 'w-24' }) => (
    <Skeleton className={`h-10 ${width} rounded-lg`} />
);

// Form field skeleton
export const SkeletonFormField: React.FC = () => (
    <div className="space-y-2">
        <Skeleton className="h-4 w-24" />
        <Skeleton className="h-12 w-full rounded-lg" />
    </div>
);

export default Skeleton;
