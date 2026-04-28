import React, { useEffect, useState } from 'react';

interface ConfettiPiece {
    id: number;
    x: number;
    color: string;
    delay: number;
    rotation: number;
}

/**
 * Confetti Animation Component
 * Shows celebration confetti on campaign completion
 */
const Confetti: React.FC<{ show: boolean; onComplete?: () => void }> = ({ show, onComplete }) => {
    const [pieces, setPieces] = useState<ConfettiPiece[]>([]);

    useEffect(() => {
        if (show) {
            const colors = [
                '#667eea', '#764ba2', '#10b981', '#f59e0b',
                '#ef4444', '#06b6d4', '#8b5cf6', '#ec4899'
            ];

            const newPieces = Array.from({ length: 50 }, (_, i) => ({
                id: i,
                x: Math.random() * 100,
                color: colors[Math.floor(Math.random() * colors.length)],
                delay: Math.random() * 0.5,
                rotation: Math.random() * 360,
            }));

            setPieces(newPieces);

            // Clean up after animation
            const timer = setTimeout(() => {
                setPieces([]);
                onComplete?.();
            }, 3000);

            return () => clearTimeout(timer);
        }
    }, [show, onComplete]);

    if (!show || pieces.length === 0) return null;

    return (
        <div className="fixed inset-0 pointer-events-none z-50 overflow-hidden">
            {pieces.map((piece) => (
                <div
                    key={piece.id}
                    className="absolute w-3 h-3 rounded-sm animate-confetti"
                    style={{
                        left: `${piece.x}%`,
                        top: '-20px',
                        backgroundColor: piece.color,
                        animationDelay: `${piece.delay}s`,
                        transform: `rotate(${piece.rotation}deg)`,
                    }}
                />
            ))}
        </div>
    );
};

export default Confetti;
