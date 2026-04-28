
import React from 'react';

interface StepIndicatorProps {
  currentStep: number;
  totalSteps: number;
}

const steps = [
  { number: 1, title: 'Connect', subtitle: 'Gmail', gradient: 'linear-gradient(135deg, #6366f1, #a855f7)', glow: 'rgba(99, 102, 241, 0.4)', color: '#6366f1' },
  { number: 2, title: 'Upload', subtitle: 'Recipients', gradient: 'linear-gradient(135deg, #10b981, #14b8a6)', glow: 'rgba(20, 184, 166, 0.4)', color: '#14b8a6' },
  { number: 3, title: 'Compose', subtitle: 'Email', gradient: 'linear-gradient(135deg, #f43f5e, #f97316)', glow: 'rgba(244, 63, 94, 0.4)', color: '#f43f5e' },
  { number: 4, title: 'Review', subtitle: '& Send', gradient: 'linear-gradient(135deg, #06b6d4, #3b82f6)', glow: 'rgba(6, 182, 212, 0.4)', color: '#06b6d4' },
];

const StepIndicator: React.FC<StepIndicatorProps> = ({ currentStep }) => {
  return (
    <div className="glass-card p-4 sm:p-5">
      {/* Desktop: horizontal layout */}
      <div className="hidden sm:flex items-center justify-between">
        {steps.map((step, idx) => (
          <React.Fragment key={step.number}>
            {/* Step Circle + Label */}
            <div className="flex flex-col items-center" style={{ minWidth: '72px' }}>
              <div
                className="w-11 h-11 rounded-xl flex items-center justify-center text-sm font-bold transition-all duration-500 relative"
                style={{
                  background: currentStep > step.number
                    ? 'linear-gradient(135deg, #14b8a6, #10b981)'
                    : currentStep === step.number
                      ? step.gradient
                      : 'rgba(148, 163, 184, 0.08)',
                  border: currentStep >= step.number ? 'none' : '1px solid rgba(148, 163, 184, 0.15)',
                  boxShadow: currentStep === step.number ? `0 0 25px ${step.glow}` : 'none',
                  color: currentStep >= step.number ? '#fff' : '#64748b',
                }}
              >
                {currentStep > step.number ? (
                  <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                    <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M5 13l4 4L19 7" />
                  </svg>
                ) : (
                  <span>{step.number}</span>
                )}

                {/* Pulse ring on current step */}
                {currentStep === step.number && (
                  <div
                    className="absolute inset-0 rounded-xl animate-ping opacity-20"
                    style={{ background: step.gradient }}
                  />
                )}
              </div>

              {/* Label */}
              <div className="mt-2 text-center">
                <span
                  className="block text-xs font-semibold transition-colors duration-300"
                  style={{ color: currentStep >= step.number ? '#f1f5f9' : '#64748b' }}
                >
                  {step.title}
                </span>
                <span
                  className="block text-[10px] transition-colors duration-300"
                  style={{ color: currentStep >= step.number ? '#94a3b8' : '#475569' }}
                >
                  {step.subtitle}
                </span>
              </div>
            </div>

            {/* Connector Line */}
            {idx < steps.length - 1 && (
              <div
                className="flex-1 mx-3 h-[2px] rounded-full overflow-hidden"
                style={{ background: 'rgba(148, 163, 184, 0.1)' }}
              >
                <div
                  className="h-full transition-all duration-700 rounded-full"
                  style={{
                    width: currentStep > step.number ? '100%' : '0%',
                    background: `linear-gradient(90deg, ${step.color}, ${steps[idx + 1].color})`,
                  }}
                />
              </div>
            )}
          </React.Fragment>
        ))}
      </div>

      {/* Mobile: compact horizontal dots + current label */}
      <div className="sm:hidden">
        <div className="flex items-center justify-center gap-3 mb-2">
          {steps.map((step, idx) => (
            <React.Fragment key={step.number}>
              <div
                className="w-8 h-8 rounded-lg flex items-center justify-center text-xs font-bold transition-all duration-300"
                style={{
                  background: currentStep > step.number
                    ? 'linear-gradient(135deg, #14b8a6, #10b981)'
                    : currentStep === step.number
                      ? step.gradient
                      : 'rgba(148, 163, 184, 0.08)',
                  border: currentStep >= step.number ? 'none' : '1px solid rgba(148, 163, 184, 0.15)',
                  color: currentStep >= step.number ? '#fff' : '#64748b',
                  boxShadow: currentStep === step.number ? `0 0 15px ${step.glow}` : 'none',
                }}
              >
                {currentStep > step.number ? '✓' : step.number}
              </div>
              {idx < steps.length - 1 && (
                <div
                  className="w-6 h-[2px] rounded-full"
                  style={{
                    background: currentStep > step.number ? step.color : 'rgba(148, 163, 184, 0.1)',
                  }}
                />
              )}
            </React.Fragment>
          ))}
        </div>
        <p className="text-center text-xs font-medium" style={{ color: '#94a3b8' }}>
          Step {currentStep}: {steps[currentStep - 1]?.title} {steps[currentStep - 1]?.subtitle}
        </p>
      </div>
    </div>
  );
};

export default StepIndicator;
