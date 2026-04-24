import React, { createContext, useContext, useState, useCallback } from 'react';

// Toast types
type ToastType = 'success' | 'error' | 'warning' | 'info';

interface Toast {
    id: string;
    type: ToastType;
    title: string;
    message?: string;
    duration?: number;
}

interface ToastContextType {
    toasts: Toast[];
    addToast: (toast: Omit<Toast, 'id'>) => void;
    removeToast: (id: string) => void;
    success: (title: string, message?: string) => void;
    error: (title: string, message?: string) => void;
    warning: (title: string, message?: string) => void;
    info: (title: string, message?: string) => void;
}

const ToastContext = createContext<ToastContextType | undefined>(undefined);

export const useToast = () => {
    const context = useContext(ToastContext);
    if (!context) {
        throw new Error('useToast must be used within a ToastProvider');
    }
    return context;
};

// Toast Provider Component
export const ToastProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    const [toasts, setToasts] = useState<Toast[]>([]);

    const removeToast = useCallback((id: string) => {
        setToasts(prev => prev.filter(t => t.id !== id));
    }, []);

    const addToast = useCallback((toast: Omit<Toast, 'id'>) => {
        const id = Date.now().toString();
        const newToast = { ...toast, id };
        setToasts(prev => [...prev, newToast]);

        // Auto remove after duration
        setTimeout(() => {
            removeToast(id);
        }, toast.duration || 4000);
    }, [removeToast]);

    const success = useCallback((title: string, message?: string) => {
        addToast({ type: 'success', title, message });
    }, [addToast]);

    const error = useCallback((title: string, message?: string) => {
        addToast({ type: 'error', title, message, duration: 6000 });
    }, [addToast]);

    const warning = useCallback((title: string, message?: string) => {
        addToast({ type: 'warning', title, message });
    }, [addToast]);

    const info = useCallback((title: string, message?: string) => {
        addToast({ type: 'info', title, message });
    }, [addToast]);

    return (
        <ToastContext.Provider value={{ toasts, addToast, removeToast, success, error, warning, info }}>
            {children}
            <ToastContainer toasts={toasts} removeToast={removeToast} />
        </ToastContext.Provider>
    );
};

// Toast Container Component
const ToastContainer: React.FC<{ toasts: Toast[]; removeToast: (id: string) => void }> = ({ toasts, removeToast }) => {
    return (
        <div className="fixed top-4 right-4 z-50 flex flex-col gap-3 max-w-sm w-full pointer-events-none">
            {toasts.map(toast => (
                <ToastItem key={toast.id} toast={toast} onClose={() => removeToast(toast.id)} />
            ))}
        </div>
    );
};

// Individual Toast Item
const ToastItem: React.FC<{ toast: Toast; onClose: () => void }> = ({ toast, onClose }) => {
    const icons: Record<ToastType, React.ReactElement> = {
        success: (
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M5 13l4 4L19 7" />
            </svg>
        ),
        error: (
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
            </svg>
        ),
        warning: (
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
            </svg>
        ),
        info: (
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
            </svg>
        ),
    };

    const colors: Record<ToastType, { gradient: string; border: string; iconBg: string; bg: string }> = {
        success: { gradient: 'linear-gradient(135deg, #14b8a6, #10b981)', border: '#14b8a6', iconBg: 'linear-gradient(135deg, #14b8a6, #10b981)', bg: 'rgba(20, 184, 166, 0.1)' },
        error: { gradient: 'linear-gradient(135deg, #f43f5e, #e11d48)', border: '#f43f5e', iconBg: 'linear-gradient(135deg, #f43f5e, #e11d48)', bg: 'rgba(244, 63, 94, 0.1)' },
        warning: { gradient: 'linear-gradient(135deg, #f59e0b, #f97316)', border: '#f59e0b', iconBg: 'linear-gradient(135deg, #f59e0b, #f97316)', bg: 'rgba(245, 158, 11, 0.1)' },
        info: { gradient: 'linear-gradient(135deg, #6366f1, #818cf8)', border: '#6366f1', iconBg: 'linear-gradient(135deg, #6366f1, #818cf8)', bg: 'rgba(99, 102, 241, 0.1)' },
    };

    const c = colors[toast.type];

    return (
        <div
            className="pointer-events-auto p-4 flex items-start gap-3 animate-slide-in rounded-xl"
            style={{
                background: c.bg,
                borderLeft: `3px solid ${c.border}`,
                backdropFilter: 'blur(24px)',
                border: `1px solid rgba(148, 163, 184, 0.12)`,
                borderLeftColor: c.border,
                borderLeftWidth: '3px',
            }}
        >
            <div className="p-2 rounded-lg text-white flex-shrink-0" style={{ background: c.iconBg }}>
                {icons[toast.type]}
            </div>
            <div className="flex-1 min-w-0">
                <p className="font-semibold text-sm" style={{ color: '#f1f5f9' }}>{toast.title}</p>
                {toast.message && (
                    <p className="text-sm mt-1" style={{ color: '#94a3b8' }}>{toast.message}</p>
                )}
            </div>
            <button
                type="button"
                onClick={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    onClose();
                }}
                className="p-2 -m-2 rounded-lg transition-all cursor-pointer flex-shrink-0"
                style={{ color: '#64748b' }}
                aria-label="Close notification"
            >
                <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
            </button>
        </div>
    );
};

export default ToastProvider;
