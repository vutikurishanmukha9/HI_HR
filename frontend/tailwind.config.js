/** @type {import('tailwindcss').Config} */
export default {
    content: [
        "./index.html",
        "./*.{js,ts,jsx,tsx}",
        "./components/**/*.{js,ts,jsx,tsx}",
        "./contexts/**/*.{js,ts,jsx,tsx}",
        "./hooks/**/*.{js,ts,jsx,tsx}",
        "./pages/**/*.{js,ts,jsx,tsx}",
        "./services/**/*.{js,ts,jsx,tsx}",
    ],
    theme: {
        extend: {
            fontFamily: {
                sans: ['Inter', '-apple-system', 'BlinkMacSystemFont', 'Segoe UI', 'Roboto', 'sans-serif'],
                display: ['Inter', '-apple-system', 'BlinkMacSystemFont', 'Segoe UI', 'Roboto', 'sans-serif'],
                mono: ['JetBrains Mono', 'ui-monospace', 'monospace'],
            },
            colors: {
                canvas: 'var(--surface-canvas)',
                card: 'var(--surface-card)',
                subdued: 'var(--surface-subdued)',
                ink: {
                    DEFAULT: 'var(--ink-primary)',
                    secondary: 'var(--ink-secondary)',
                    muted: 'var(--ink-muted)',
                },
                rule: {
                    DEFAULT: 'var(--border-default)',
                    subtle: 'var(--border-subtle)',
                    strong: 'var(--border-strong)',
                },
                accent: {
                    DEFAULT: 'var(--accent-primary)',
                    hover: 'var(--accent-hover)',
                    active: 'var(--accent-active)',
                    subtle: 'var(--accent-subtle)',
                    light: 'var(--accent-subtle)',
                    border: 'var(--accent-border)',
                    text: 'var(--accent-text)',
                },
                success: {
                    DEFAULT: 'var(--success-primary)',
                    subtle: 'var(--success-subtle)',
                    light: 'var(--success-subtle)',
                    border: 'var(--success-border)',
                    text: 'var(--success-text)',
                },
                error: {
                    DEFAULT: 'var(--error-primary)',
                    subtle: 'var(--error-subtle)',
                    light: 'var(--error-subtle)',
                    border: 'var(--error-border)',
                    text: 'var(--error-text)',
                },
                warning: {
                    DEFAULT: 'var(--warning-primary)',
                    subtle: 'var(--warning-subtle)',
                    light: 'var(--warning-subtle)',
                    border: 'var(--warning-border)',
                    text: 'var(--warning-text)',
                },
                info: {
                    DEFAULT: 'var(--info-primary)',
                    subtle: 'var(--info-subtle)',
                    light: 'var(--info-subtle)',
                    border: 'var(--info-border)',
                    text: 'var(--info-text)',
                },
                border: 'var(--border-default)',
            },
            boxShadow: {
                xs: 'var(--shadow-xs)',
                sm: 'var(--shadow-sm)',
                md: 'var(--shadow-md)',
                lg: 'var(--shadow-lg)',
                xl: 'var(--shadow-xl)',
            },
            borderRadius: {
                xs: 'var(--radius-xs)',
                sm: 'var(--radius-sm)',
                md: 'var(--radius-md)',
                lg: 'var(--radius-lg)',
                xl: 'var(--radius-xl)',
                '2xl': 'var(--radius-2xl)',
            },
            animation: {
                'fade-in': 'fadeIn 200ms cubic-bezier(0.16, 1, 0.3, 1) forwards',
                'slide-in': 'slideIn 250ms cubic-bezier(0.16, 1, 0.3, 1) forwards',
                'slide-down': 'slideDown 200ms cubic-bezier(0.16, 1, 0.3, 1) forwards',
            },
            keyframes: {
                fadeIn: {
                    '0%': { opacity: '0', transform: 'translateY(4px)' },
                    '100%': { opacity: '1', transform: 'translateY(0)' },
                },
                slideIn: {
                    '0%': { opacity: '0', transform: 'translateX(16px)' },
                    '100%': { opacity: '1', transform: 'translateX(0)' },
                },
                slideDown: {
                    '0%': { opacity: '0', transform: 'translateY(-8px)' },
                    '100%': { opacity: '1', transform: 'translateY(0)' },
                },
            },
        },
    },
    plugins: [],
};
