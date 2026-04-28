import { Component, ErrorInfo, ReactNode } from 'react';

interface ErrorBoundaryProps {
    children: ReactNode;
    fallback?: ReactNode;
}

interface ErrorBoundaryState {
    hasError: boolean;
    error: Error | null;
}

class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
    public state: ErrorBoundaryState = {
        hasError: false,
        error: null,
    };

    public static getDerivedStateFromError(error: Error): ErrorBoundaryState {
        return { hasError: true, error };
    }

    public componentDidCatch(error: Error, errorInfo: ErrorInfo): void {
        // Log error in development
        if (import.meta.env.DEV) {
            console.error('Error caught by boundary:', error, errorInfo);
        }
    }

    private handleRetry = (): void => {
        this.setState({ hasError: false, error: null });
    };

    public render(): ReactNode {
        if (this.state.hasError) {
            if (this.props.fallback) {
                return this.props.fallback;
            }

            return (
                <div className="min-h-screen flex items-center justify-center" style={{ background: 'linear-gradient(145deg, #0c0e1a 0%, #111327 40%, #141625 100%)' }}>
                    <div className="text-center p-8 rounded-2xl max-w-md" style={{ background: 'rgba(22, 24, 42, 0.85)', border: '1px solid rgba(148, 163, 184, 0.12)', backdropFilter: 'blur(24px)' }}>
                        <div className="inline-flex items-center justify-center w-16 h-16 rounded-full mb-6" style={{ background: 'rgba(244, 63, 94, 0.15)' }}>
                            <svg className="w-8 h-8" style={{ color: '#fb7185' }} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
                            </svg>
                        </div>
                        <h2 className="text-2xl font-bold mb-3" style={{ color: '#f1f5f9' }}>Something went wrong</h2>
                        <p className="mb-6" style={{ color: '#94a3b8', fontSize: '0.9375rem' }}>
                            An unexpected error occurred. Please try again or refresh the page.
                        </p>
                        {this.state.error && (
                            <details className="text-left mb-6 p-3 rounded-lg" style={{ background: 'rgba(0,0,0,0.2)' }}>
                                <summary className="text-sm cursor-pointer" style={{ color: '#64748b' }}>Error details</summary>
                                <pre className="text-xs mt-2 overflow-auto" style={{ color: '#fb7185' }}>
                                    {this.state.error.message}
                                </pre>
                            </details>
                        )}
                        <div className="flex gap-3 justify-center">
                            <button
                                onClick={this.handleRetry}
                                className="px-6 py-3 rounded-xl font-semibold text-white text-sm transition-all hover:scale-[1.02] active:scale-[0.98]"
                                style={{ background: 'linear-gradient(135deg, #6366f1, #a855f7)', boxShadow: '0 4px 15px rgba(99, 102, 241, 0.35)' }}
                            >
                                Try Again
                            </button>
                            <button
                                onClick={() => window.location.reload()}
                                className="px-6 py-3 rounded-xl font-medium text-sm transition-all"
                                style={{ border: '1px solid rgba(148, 163, 184, 0.15)', color: '#94a3b8' }}
                            >
                                Refresh Page
                            </button>
                        </div>
                    </div>
                </div>
            );
        }

        return this.props.children;
    }
}

export default ErrorBoundary;
