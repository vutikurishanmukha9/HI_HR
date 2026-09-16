import React, { useState } from 'react';
import { Credentials } from '../types';

interface CredentialsFormProps {
    onSave: (credentials: Credentials) => Promise<{ success: boolean; error?: string } | void> | void;
    initialCredentials?: Credentials | null;
}

const CredentialsForm: React.FC<CredentialsFormProps> = ({ onSave, initialCredentials }) => {
    const [email, setEmail] = useState(initialCredentials?.email || '');
    const [appPassword, setAppPassword] = useState(initialCredentials?.appPassword || '');
    const [error, setError] = useState('');
    const [showPassword, setShowPassword] = useState(false);
    const [isVerifying, setIsVerifying] = useState(false);

    const handleSubmit = async (e: React.FormEvent) => {
        e.preventDefault();

        const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
        if (!email.trim() || !emailRegex.test(email.trim())) {
            setError('Please enter a valid email address.');
            return;
        }

        const cleanPassword = appPassword.replace(/\s/g, '');
        if (cleanPassword.length !== 16) {
            setError('Google App Password must be exactly 16 characters (spaces are ignored).');
            return;
        }

        setError('');
        setIsVerifying(true);
        try {
            const res = await onSave({ email: email.trim(), appPassword: cleanPassword });
            if (res && !res.success && res.error) {
                setError(res.error);
            }
        } catch (err: any) {
            setError(err.message || 'Verification failed');
        } finally {
            setIsVerifying(false);
        }
    };

    return (
        <div className="max-w-xl mx-auto space-y-6">
            {/* Header */}
            <div className="text-center mb-6">
                <div className="inline-flex items-center justify-center w-12 h-12 rounded-xl bg-accent-light border border-neutral-200 text-accent mb-3 shadow-xs">
                    <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 7a2 2 0 012 2m4 0a6 6 0 01-7.743 5.743L11 17H9v2H7v2H4a1 1 0 01-1-1v-2.586a1 1 0 01.293-.707l5.964-5.964A6 6 0 1121 9z" />
                    </svg>
                </div>
                <h2 className="text-xl font-bold text-ink">Connect Your Gmail Sender</h2>
                <p className="text-sm text-ink-muted mt-1">Enter your sender email and Google App Password</p>
            </div>

            {/* Form Card */}
            <div className="surface-card p-6 border border-border shadow-xs space-y-5 animate-fadeIn">
                {/* Security Notice */}
                <div className="p-3.5 rounded-xl bg-amber-50 border border-amber-200">
                    <div className="flex items-start gap-3">
                        <svg className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" viewBox="0 0 20 20" fill="currentColor">
                            <path fillRule="evenodd" d="M8.485 2.495c.673-1.167 2.357-1.167 3.03 0l6.28 10.875c.673 1.167-.17 2.625-1.516 2.625H3.72c-1.347 0-2.189-1.458-1.515-2.625L8.485 2.495zM10 5a.75.75 0 01.75.75v3.5a.75.75 0 01-1.5 0v-3.5A.75.75 0 0110 5zm0 9a1 1 0 100-2 1 1 0 000 2z" clipRule="evenodd" />
                        </svg>
                        <p className="text-xs leading-relaxed text-amber-900">
                            <span className="font-semibold">Security Requirement:</span> Do not use your regular account password. Generate a 16-character dedicated App Password from your Google Account.
                            <a
                                href="https://support.google.com/accounts/answer/185833"
                                target="_blank"
                                rel="noopener noreferrer"
                                className="ml-1 text-amber-700 underline font-medium hover:text-amber-900"
                            >
                                Setup instructions
                            </a>
                        </p>
                    </div>
                </div>

                {/* Form */}
                <form onSubmit={handleSubmit} className="space-y-4">
                    {/* Email Field */}
                    <div className="space-y-1.5">
                        <label htmlFor="email" className="block text-xs font-semibold uppercase tracking-wider text-ink-secondary">
                            Sender Email Address
                        </label>
                        <div className="relative">
                            <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-ink-muted">
                                <svg className="w-4 h-4 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z" />
                                </svg>
                            </div>
                            <input
                                id="email"
                                type="email"
                                value={email}
                                onChange={(e) => setEmail(e.target.value)}
                                placeholder="you@gmail.com"
                                className="input input-icon-left !pl-11"
                                style={{ paddingLeft: '2.75rem' }}
                                required
                            />
                        </div>
                    </div>

                    {/* Password Field */}
                    <div className="space-y-1.5">
                        <label htmlFor="appPassword" className="block text-xs font-semibold uppercase tracking-wider text-ink-secondary">
                            Google App Password
                        </label>
                        <div className="relative">
                            <div className="absolute inset-y-0 left-0 pl-3.5 flex items-center pointer-events-none text-ink-muted">
                                <svg className="w-4 h-4 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" />
                                </svg>
                            </div>
                            <input
                                id="appPassword"
                                type={showPassword ? "text" : "password"}
                                value={appPassword}
                                onChange={(e) => setAppPassword(e.target.value)}
                                placeholder="xxxx xxxx xxxx xxxx"
                                className="input input-icon-both !pl-11 !pr-11 font-mono"
                                style={{ paddingLeft: '2.75rem', paddingRight: '2.75rem' }}
                                required
                            />
                            <button
                                type="button"
                                onClick={() => setShowPassword(!showPassword)}
                                className="absolute inset-y-0 right-0 pr-3.5 flex items-center text-ink-muted hover:text-ink transition-colors"
                            >
                                {showPassword ? (
                                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l3.59 3.59m0 0A9.953 9.953 0 0112 5c4.478 0 8.268 2.943 9.543 7a10.025 10.025 0 01-4.132 5.411m0 0L21 21" />
                                    </svg>
                                ) : (
                                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" />
                                        <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" />
                                    </svg>
                                )}
                            </button>
                        </div>
                        <p className="text-[11px] text-ink-muted">16 characters without spaces.</p>
                    </div>

                    {/* Error Message */}
                    {error && (
                        <div className="p-3 rounded-lg bg-danger-light border border-red-200 text-danger animate-fadeIn">
                            <p className="text-xs font-medium flex items-center gap-2">
                                <svg className="w-4 h-4 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                                    <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
                                </svg>
                                {error}
                            </p>
                        </div>
                    )}

                    {/* Submit Button */}
                    <div className="flex justify-end pt-2">
                        <button
                            type="submit"
                            disabled={!email || !appPassword || isVerifying}
                            className="btn btn-primary px-8 py-2.5 text-xs font-semibold shadow-xs flex items-center gap-2"
                        >
                            {isVerifying ? (
                                <>
                                    <div className="w-3.5 h-3.5 border-2 border-white border-t-transparent rounded-full animate-spin" />
                                    <span>Verifying with Gmail SMTP...</span>
                                </>
                            ) : (
                                <span>Save &amp; Continue &rarr;</span>
                            )}
                        </button>
                    </div>
                </form>
            </div>
        </div>
    );
};

export default CredentialsForm;
