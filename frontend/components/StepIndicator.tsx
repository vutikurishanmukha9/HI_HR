import React from 'react';

interface StepIndicatorProps {
  currentStep: number;
  totalSteps: number;
}

const steps = [
  { number: 1, title: 'Connect', subtitle: 'Gmail' },
  { number: 2, title: 'Upload', subtitle: 'Recipients' },
  { number: 3, title: 'Compose', subtitle: 'Email' },
  { number: 4, title: 'Review', subtitle: '& Send' },
];

const StepIndicator: React.FC<StepIndicatorProps> = ({ currentStep }) => {
  return (
    <div className="surface-card p-4 sm:p-5">
      {/* Desktop layout */}
      <div className="hidden sm:block relative">
        {/* Connector track positioned at the exact vertical center of the 40px badges (top = 20px) */}
        <div className="absolute top-5 left-14 right-14 -translate-y-1/2 h-[2px] bg-neutral-200" aria-hidden="true">
          <div
            className="h-full bg-accent transition-all duration-300"
            style={{
              width: `${((Math.min(currentStep, steps.length) - 1) / (steps.length - 1)) * 100}%`,
            }}
          />
        </div>

        {/* Step columns */}
        <div className="relative z-10 flex justify-between">
          {steps.map((step) => {
            const isCompleted = currentStep > step.number;
            const isActive = currentStep === step.number;

            return (
              <div key={step.number} className="flex flex-col items-center w-28 text-center">
                {/* Step Badge */}
                <div
                  className={`w-10 h-10 rounded-none flex items-center justify-center text-sm font-bold transition-all duration-200 ${
                    isCompleted
                      ? 'bg-accent text-white ring-4 ring-white shadow-xs'
                      : isActive
                      ? 'bg-accent text-white ring-4 ring-neutral-200 shadow-xs'
                      : 'bg-white text-ink-muted border border-border ring-4 ring-white'
                  }`}
                >
                  {isCompleted ? (
                    <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M5 13l4 4L19 7" />
                    </svg>
                  ) : (
                    <span>{step.number}</span>
                  )}
                </div>

                {/* Label & Subtitle safely positioned below with no overlapping */}
                <div className="mt-2.5 space-y-0.5">
                  <span
                    className={`block text-xs font-semibold leading-tight ${
                      isActive || isCompleted ? 'text-ink' : 'text-ink-muted'
                    }`}
                  >
                    {step.title}
                  </span>
                  <span
                    className={`block text-[11px] leading-tight ${
                      isActive || isCompleted ? 'text-ink-secondary' : 'text-ink-muted'
                    }`}
                  >
                    {step.subtitle}
                  </span>
                </div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Mobile: compact indicator */}
      <div className="sm:hidden">
        <div className="flex items-center justify-center gap-2 mb-2">
          {steps.map((step, idx) => {
            const isCompleted = currentStep > step.number;
            const isActive = currentStep === step.number;

            return (
              <React.Fragment key={step.number}>
                <div
                  className={`w-8 h-8 rounded-none flex items-center justify-center text-xs font-bold transition-all ${
                    isCompleted
                      ? 'bg-accent text-white'
                      : isActive
                      ? 'bg-accent text-white ring-2 ring-neutral-200'
                      : 'bg-neutral-100 text-ink-muted border border-border'
                  }`}
                >
                  {isCompleted ? (
                    <svg className="w-4 h-4 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                      <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M5 13l4 4L19 7" />
                    </svg>
                  ) : (
                    <span>{step.number}</span>
                  )}
                </div>
                {idx < steps.length - 1 && (
                  <div
                    className={`w-6 h-[2px] rounded-none transition-colors ${
                      currentStep > step.number ? 'bg-accent' : 'bg-neutral-200'
                    }`}
                  />
                )}
              </React.Fragment>
            );
          })}
        </div>
        <p className="text-center text-xs font-medium text-ink-secondary">
          Step {currentStep} of 4: <span className="font-semibold text-ink">{steps[currentStep - 1]?.title} {steps[currentStep - 1]?.subtitle}</span>
        </p>
      </div>
    </div>
  );
};

export default StepIndicator;
