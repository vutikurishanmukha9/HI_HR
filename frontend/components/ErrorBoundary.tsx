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
                <div className="min-h-screen flex items-center justify-center p-4 bg-canvas">
                    <div className="surface-card text-center p-8 max-w-md border border-border shadow-md">
                        <div className="inline-flex items-center justify-center w-14 h-14 rounded-2xl mb-4 bg-danger-light border border-red-200 text-danger shadow-xs">
                            <svg
                                className="w-7 h-7"
                                fill="none"
                                stroke="currentColor"
                                viewBox="0 0 24 24"
                            >
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
                            </svg>
                        </div>
                        <h2 className="text-xl font-bold mb-1.5 text-ink">
                            An unexpected error occurred
                        </h2>
                        <p className="text-xs text-ink-muted mb-5 leading-relaxed">
                            Something interrupted this operation. You can retry or reload your workspace.
                        </p>
                        {this.state.error && (
                            <details open className="text-left mb-5 p-3 rounded-lg bg-neutral-50 border border-border">
                                <summary className="text-xs font-semibold text-ink-secondary cursor-pointer">
                                    Technical error details
                                </summary>
                                <pre className="text-[11px] mt-2 overflow-auto text-danger font-mono leading-tight whitespace-pre-wrap">
                                    {this.state.error.message}
                                </pre>
                                {this.state.error.stack && (
                                    <pre className="text-[10px] mt-2 overflow-auto text-ink-muted font-mono leading-tight max-h-32">
                                        {this.state.error.stack}
                                    </pre>
                                )}
                            </details>
                        )}
                        <div className="flex gap-2 justify-center">
                            <button
                                onClick={this.handleRetry}
                                className="btn btn-primary text-xs px-4 py-2 shadow-xs"
                            >
                                Try Again
                            </button>
                            <button
                                onClick={() => window.location.reload()}
                                className="btn btn-secondary text-xs px-4 py-2"
                            >
                                Reload Workspace
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
