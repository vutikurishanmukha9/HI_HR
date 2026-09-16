import React from 'react';

interface StatCardProps {
    title: string;
    value: string | number;
    subtitle?: string;
    icon: React.ReactNode;
}

/**
 * StatCard Component
 * Displays a single metric card using the Hallmark design system.
 */
const StatCard: React.FC<StatCardProps> = ({ 
    title, 
    value, 
    subtitle, 
    icon
}) => {
    return (
        <div className="surface-card p-5 relative overflow-hidden border border-border shadow-xs">
            <div className="flex items-start justify-between gap-3">
                <div className="min-w-0 flex-1">
                    <p className="text-xs font-semibold uppercase tracking-wider text-ink-secondary truncate">
                        {title}
                    </p>
                    <h3 className="text-2xl font-bold mt-1.5 mb-0.5 text-ink tabular-nums leading-tight truncate">
                        {value}
                    </h3>
                    {subtitle && (
                        <p className="text-xs text-ink-muted truncate">
                            {subtitle}
                        </p>
                    )}
                </div>
                
                <div className="w-10 h-10 rounded-xl flex items-center justify-center bg-accent-light text-accent border border-neutral-200 flex-shrink-0">
                    {icon}
                </div>
            </div>
        </div>
    );
};

export default StatCard;
