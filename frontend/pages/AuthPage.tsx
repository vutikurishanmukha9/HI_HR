import React, { useState } from 'react';
import { useAuth } from '../contexts/AuthContext';
import PasswordStrengthMeter from '../components/PasswordStrengthMeter';

/**
 * AuthPage Component - Premium Login/Register Design
 * Handles user authentication with animated background
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
        <div className="min-h-screen flex items-center justify-center p-4">
            {/* Animated background orbs — warmer tones */}
            <div className="fixed inset-0 overflow-hidden pointer-events-none">
                <div className="absolute top-1/4 left-1/4 w-80 h-80 rounded-full blur-[100px] animate-pulse"
                    style={{ background: 'rgba(99, 102, 241, 0.12)' }} />
                <div className="absolute bottom-1/3 right-1/4 w-72 h-72 rounded-full blur-[100px] animate-pulse"
                    style={{ background: 'rgba(168, 85, 247, 0.1)', animationDelay: '1s' }} />
                <div className="absolute top-1/2 right-1/3 w-64 h-64 rounded-full blur-[100px] animate-pulse"
                    style={{ background: 'rgba(20, 184, 166, 0.08)', animationDelay: '2s' }} />
            </div>

            {/* Main Card */}
            <div className="relative w-full max-w-md fade-in">
                <div className="glass-card p-8 md:p-10">
                    {/* Logo & Title */}
                    <div className="text-center mb-8">
                        <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl mb-5 shadow-lg"
                            style={{ background: 'linear-gradient(135deg, #6366f1 0%, #a855f7 50%, #06b6d4 100%)' }}>
                            <span className="text-white font-extrabold text-2xl" style={{ fontFamily: "'Outfit', sans-serif" }}>Hi</span>
                        </div>
                        <h1 className="text-4xl font-extrabold mb-2 text-brand">HiHR</h1>
                        <p style={{ color: '#94a3b8', fontSize: '0.9375rem' }}>Smart HR Email Outreach</p>
                    </div>

                    {/* Tab Switcher */}
                    <div className="flex p-1 mb-8 rounded-xl" style={{ background: 'rgba(148, 163, 184, 0.06)' }}>
                        <button
                            onClick={() => setIsLogin(true)}
                            className={`flex-1 py-3 px-4 rounded-lg font-semibold transition-all duration-300 text-sm ${isLogin
                                ? 'text-white shadow-lg'
                                : 'hover:text-white'
                                }`}
                            style={{
                                background: isLogin ? 'linear-gradient(135deg, #6366f1, #a855f7)' : 'transparent',
                                color: isLogin ? '#fff' : '#64748b',
                            }}
                        >
                            Sign In
                        </button>
                        <button
                            onClick={() => setIsLogin(false)}
                            className={`flex-1 py-3 px-4 rounded-lg font-semibold transition-all duration-300 text-sm ${!isLogin
                                ? 'text-white shadow-lg'
                                : 'hover:text-white'
                                }`}
                            style={{
                                background: !isLogin ? 'linear-gradient(135deg, #6366f1, #a855f7)' : 'transparent',
                                color: !isLogin ? '#fff' : '#64748b',
                            }}
                        >
                            Register
                        </button>
                    </div>

                    {/* Form */}
                    <form onSubmit={handleSubmit} className="space-y-5">
                        {!isLogin && (
                            <div className="slide-up">
                                <label className="block text-sm font-medium mb-2" style={{ color: '#94a3b8', letterSpacing: '0.03em' }}>
                                    Full Name
                                </label>
                                <input
                                    type="text"
                                    value={fullName}
                                    onChange={(e) => setFullName(e.target.value)}
                                    placeholder="John Doe"
                                    required={!isLogin}
                                />
                            </div>
                        )}

                        <div>
                            <label className="block text-sm font-medium mb-2" style={{ color: '#94a3b8', letterSpacing: '0.03em' }}>
                                Email Address
                            </label>
                            <input
                                type="email"
                                value={email}
                                onChange={(e) => setEmail(e.target.value)}
                                placeholder="you@example.com"
                                required
                            />
                        </div>

                        <div>
                            <label className="block text-sm font-medium mb-2" style={{ color: '#94a3b8', letterSpacing: '0.03em' }}>
                                Password
                            </label>
                            <input
                                type="password"
                                value={password}
                                onChange={(e) => setPassword(e.target.value)}
                                placeholder="••••••••"
                                required
                                minLength={8}
                            />
                            {!isLogin && <PasswordStrengthMeter password={password} />}
                        </div>

                        {error && (
                            <div className="p-4 rounded-xl" style={{ background: 'rgba(244, 63, 94, 0.1)', border: '1px solid rgba(244, 63, 94, 0.25)' }}>
                                <p className="text-sm flex items-center gap-2" style={{ color: '#fb7185' }}>
                                    <svg className="w-4 h-4 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                                        <path fillRule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clipRule="evenodd" />
                                    </svg>
                                    {error}
                                </p>
                            </div>
                        )}

                        <button
                            type="submit"
                            disabled={loading}
                            className="w-full py-4 rounded-xl font-semibold text-white transition-all duration-300 disabled:opacity-40 disabled:cursor-not-allowed text-sm"
                            style={{
                                background: 'linear-gradient(135deg, #6366f1 0%, #a855f7 100%)',
                                boxShadow: loading ? 'none' : '0 4px 20px rgba(99, 102, 241, 0.35)',
                                letterSpacing: '0.02em',
                            }}
                        >
                            {loading ? (
                                <span className="flex items-center justify-center gap-2">
                                    <svg className="animate-spin h-5 w-5" viewBox="0 0 24 24">
                                        <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none" />
                                        <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z" />
                                    </svg>
                                    Please wait...
                                </span>
                            ) : isLogin ? 'Sign In' : 'Create Account'}
                        </button>
                    </form>

                    {/* Info Box */}
                    <div className="mt-6 p-4 rounded-xl" style={{ background: 'rgba(20, 184, 166, 0.08)', border: '1px solid rgba(20, 184, 166, 0.2)' }}>
                        <p className="text-sm flex items-start gap-2" style={{ color: '#5eead4' }}>
                            <svg className="w-5 h-5 mt-0.5 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20">
                                <path fillRule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clipRule="evenodd" />
                            </svg>
                            <span>Your Gmail credentials are encrypted end-to-end and stored securely on our servers.</span>
                        </p>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default AuthPage;
