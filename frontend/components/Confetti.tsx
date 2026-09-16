import React, { useEffect } from 'react';

/**
 * Confetti Animation Component
 * Shows celebration confetti on campaign completion
 *
 * NOTE: Gutted per Hallmark design system — celebratory animations are an anti-pattern.
 * Renders nothing but preserves the interface so the parent doesn't break.
 */
const Confetti: React.FC<{ show: boolean; onComplete?: () => void }> = ({ onComplete }) => {
  useEffect(() => {
    onComplete?.();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps
  return null;
};

export default Confetti;
