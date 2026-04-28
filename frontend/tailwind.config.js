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
                'display': ['Outfit', 'Inter', 'system-ui', 'sans-serif'],
                'sans': ['Inter', 'system-ui', 'sans-serif'],
            },
            colors: {
                surface: {
                    primary: '#0c0e1a',
                    secondary: '#111327',
                    tertiary: '#141625',
                    card: 'rgba(22, 24, 42, 0.85)',
                },
                accent: {
                    indigo: '#6366f1',
                    violet: '#a855f7',
                    teal: '#14b8a6',
                    emerald: '#10b981',
                    rose: '#f43f5e',
                    amber: '#f59e0b',
                    cyan: '#06b6d4',
                    blue: '#3b82f6',
                },
            },
            animation: {
                'pulse-slow': 'pulse 3s cubic-bezier(0.4, 0, 0.6, 1) infinite',
                'fade-in': 'fadeIn 0.5s ease-out forwards',
                'slide-up': 'slideUp 0.5s ease-out forwards',
                'glow': 'glowPulse 2.5s ease-in-out infinite',
            },
            keyframes: {
                fadeIn: {
                    '0%': { opacity: '0', transform: 'translateY(16px)' },
                    '100%': { opacity: '1', transform: 'translateY(0)' },
                },
                slideUp: {
                    '0%': { opacity: '0', transform: 'translateY(30px)' },
                    '100%': { opacity: '1', transform: 'translateY(0)' },
                },
                glowPulse: {
                    '0%, 100%': { boxShadow: '0 0 20px rgba(99, 102, 241, 0.2)' },
                    '50%': { boxShadow: '0 0 35px rgba(99, 102, 241, 0.35)' },
                },
            },
        },
    },
    plugins: [],
}
