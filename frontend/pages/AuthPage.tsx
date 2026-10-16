import React, { useState } from 'react';
import { useAuth } from '../contexts/AuthContext';
import PasswordStrengthMeter from '../components/PasswordStrengthMeter';

/**
 * AuthPage Component - Modern Light Design System
 * Crisp, focused authentication with strong typography and semantic feedback
 */
const AuthPage: React.FC = () => {
    const { login, register } = useAuth();
    const [isLogin, setIsLogin] = useState(true);
    const [email, setEmail] = useState('');
    const [password, setPassword] = useState('');
    const [fullName, setFullName] = useState('');
    const [error, setError] = useState('');
    const [loading, setLoading] = useState(false);

    const handleSubmit = async (e: React.FormEvent) => {
        e.preventDefault();
        setError('');
        setLoading(true);

        try {
            if (isLogin) {
                await login(email, password);
            } else {
                if (!fullName.trim()) {
                    setError('Full name is required');
                    setLoading(false);
                    return;
                }
                await register(email, password, fullName);
            }
        } catch (err: any) {
            setError(err.message || 'Authentication failed');
        } finally {
            setLoading(false);
        }
    };

    return (
        <div className="min-h-screen flex items-center justify-center p-4 sm:p-6 bg-canvas">
            <div className="w-full max-w-md">
                {/* Main Card */}
                <div className="surface-card p-8 sm:p-10 border border-border shadow-md">
                    {/* Brand Header */}
                    <div className="text-center mb-8">
                        <div className="inline-flex items-center justify-center w-12 h-12 rounded-none bg-accent text-white font-bold text-xl mb-3 shadow-sm">
                            Hi
                        </div>
                        <h1 className="text-2xl font-bold tracking-tight text-ink">
                            {isLogin ? 'Welcome back to HiHR' : 'Create your account'}
                        </h1>
                        <p className="text-sm text-ink-muted mt-1.5">
                            {isLogin ? 'Sign in to access your outreach campaigns' : 'Start sending personalized cold outreach in minutes'}
                        </p>
                    </div>

                    {/* Segmented Control */}
                    <div className="flex p-1 mb-6 rounded-none bg-canvas border border-border">
                        <button
                            type="button"
                            onClick={() => { setIsLogin(true); setError(''); }}
                            className={`flex-1 py-2.5 px-4 rounded-none text-sm font-semibold transition-all duration-150 ${
                                isLogin
                                    ? 'bg-card text-ink shadow-xs border border-border'
                                    : 'text-ink-muted hover:text-ink'
                            }`}
                        >
                            Sign In
                        </button>
                        <button
                            type="button"
                            onClick={() => { setIsLogin(false); setError(''); }}
                            className={`flex-1 py-2.5 px-4 rounded-none text-sm font-semibold transition-all duration-150 ${
                                !isLogin
                                    ? 'bg-card text-ink shadow-xs border border-border'
                                    : 'text-ink-muted hover:text-ink'
                            }`}
                        >
                            Register
                        </button>
                    </div>

                    {/* Form */}
                    <form onSubmit={handleSubmit} className="space-y-4">
                        {!isLogin && (
                            <div className="space-y-1.5 animate-fadeIn">
                                <label className="block text-xs font-semibold uppercase tracking-wider text-ink-secondary">
                                    Full Name
                                </label>
                                <input
                                    type="text"
                                    value={fullName}
                                    onChange={(e) => setFullName(e.target.value)}
                                    placeholder="Jane Doe"
                                    className="input"
                                    required={!isLogin}
                                />
                            </div>
                        )}

                        <div className="space-y-1.5">
                            <label className="block text-xs font-semibold uppercase tracking-wider text-ink-secondary">
                                Email Address
                            </label>
                            <input
                                type="email"
                                value={email}
                                onChange={(e) => setEmail(e.target.value)}
                                placeholder="name@company.com"
                                className="input"
                                required
                            />
                        </div>

                        <div className="space-y-1.5">
                            <label className="block text-xs font-semibold uppercase tracking-wider text-ink-secondary">
                                Password
                            </label>
                            <input
                                type="password"
                                value={password}
                                onChange={(e) => setPassword(e.target.value)}
                                placeholder="••••••••"
                                className="input"
                                required
                                minLength={8}
                            />
                            {!isLogin && <PasswordStrengthMeter password={password} />}
                        </div>

                        {error && (
                            <div className="p-3.5 rounded-none flex items-start gap-2.5 bg-danger-light border border-red-200 text-danger animate-fadeIn">
                                <svg className="w-5 h-5 flex-shrink-0 mt-0.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                                    <path strokeLinecap="round" strokeLinejoin="round" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z" />
                                </svg>
                                <p className="text-sm font-medium leading-snug break-words min-w-0 flex-1">{error}</p>
                            </div>
                        )}

                        <button
                            type="submit"
                            disabled={loading}
                            className="btn btn-primary w-full py-3 text-sm font-semibold mt-2 shadow-xs"
                        >
                            {loading ? (
                                <span className="flex items-center justify-center gap-2">
                                    <svg className="animate-spin h-4 w-4 text-white" viewBox="0 0 24 24" fill="none">
                                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
                                        <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                                    </svg>
                                    Processing…
                                </span>
                            ) : isLogin ? 'Sign In' : 'Create Account'}
                        </button>
                    </form>

                    {/* Security Info Box */}
                    <div className="mt-6 p-3.5 rounded-none bg-subdued border border-border flex items-start gap-2.5">
                        <svg className="w-4 h-4 text-ink-muted flex-shrink-0 mt-0.5" fill="none" viewBox="0 0 24 24" stroke="currentColor" strokeWidth={2}>
                            <path strokeLinecap="round" strokeLinejoin="round" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" />
                        </svg>
                        <p className="text-xs text-ink-secondary leading-relaxed">
                            Your Gmail App Passwords and outreach configurations are encrypted with AES-256 before storage.
                        </p>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default AuthPage;
