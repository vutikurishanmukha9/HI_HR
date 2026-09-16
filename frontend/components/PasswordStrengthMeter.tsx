import React, { useMemo } from 'react';

interface PasswordStrengthMeterProps {
    password: string;
}

interface StrengthResult {
    score: number;
    label: string;
    color: string;
    feedback: string[];
}

const PasswordStrengthMeter: React.FC<PasswordStrengthMeterProps> = ({ password }) => {
    const strength: StrengthResult = useMemo(() => {
        if (!password) {
            return { score: 0, label: '', color: 'transparent', feedback: [] };
        }

        let score = 0;
        const feedback: string[] = [];

        // Length checks
        if (password.length >= 8) {
            score += 1;
        } else {
            feedback.push('At least 8 characters');
        }

        if (password.length >= 12) {
            score += 1;
        }

        // Uppercase check
        if (/[A-Z]/.test(password)) {
            score += 1;
        } else {
            feedback.push('Add uppercase letter');
        }

        // Lowercase check
        if (/[a-z]/.test(password)) {
            score += 1;
        } else {
            feedback.push('Add lowercase letter');
        }

        // Number check
        if (/[0-9]/.test(password)) {
            score += 1;
        } else {
            feedback.push('Add a number');
        }

        // Special character check
        if (/[!@#$%^&*(),.?":{}|<>]/.test(password)) {
            score += 1;
        } else {
            feedback.push('Add special character');
        }

        // Determine label and color: weak to error, medium to warning, strong to success
        let label: string;
        let color: string;

        if (score <= 2) {
            label = 'Weak';
            color = '#ef4444';
        } else if (score <= 4) {
            label = 'Medium';
            color = '#f59e0b';
        } else {
            label = 'Strong';
            color = '#10b981';
        }

        return { score, label, color, feedback };
    }, [password]);

    if (!password) {
        return null;
    }

    const percentage = (strength.score / 6) * 100;

    return (
        <div className="mt-2.5 space-y-2">
            {/* Strength bar */}
            <div className="flex items-center gap-3">
                <div className="flex-1 h-1.5 rounded-full overflow-hidden bg-neutral-200">
                    <div
                        className="h-full rounded-full transition-all duration-300"
                        style={{
                            width: `${percentage}%`,
                            backgroundColor: strength.color,
                        }}
                    />
                </div>
                <span
                    className="text-xs font-semibold min-w-[50px] text-right"
                    style={{ color: strength.color }}
                >
                    {strength.label}
                </span>
            </div>

            {/* Feedback */}
            {strength.feedback.length > 0 && (
                <div className="flex flex-wrap gap-1">
                    {strength.feedback.slice(0, 3).map((tip, index) => (
                        <span
                            key={index}
                            className="text-[11px] px-2 py-0.5 rounded-md bg-neutral-100 text-ink-secondary border border-border font-medium"
                        >
                            {tip}
                        </span>
                    ))}
                </div>
            )}
        </div>
    );
};

export default PasswordStrengthMeter;
