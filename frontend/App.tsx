import React from 'react';
import { useAuth } from './contexts/AuthContext';
import AuthPage from './pages/AuthPage';
import DashboardPage from './pages/DashboardPage';

/**
 * Main App Component
 * Handles routing between Auth and Dashboard based on authentication state
 */
const App: React.FC = () => {
  const { isAuthenticated, isLoading } = useAuth();

  // Show loading screen while checking auth
  if (isLoading) {
    return (
      <div className="min-h-screen flex items-center justify-center" style={{ background: '#0c0e1a' }}>
        <div className="text-center">
          <div className="w-14 h-14 mx-auto mb-4 rounded-2xl flex items-center justify-center" style={{ background: 'linear-gradient(135deg, #6366f1, #a855f7)' }}>
            <span className="text-white font-extrabold text-xl" style={{ fontFamily: "'Outfit', sans-serif" }}>Hi</span>
          </div>
          <div className="animate-spin rounded-full h-8 w-8 border-2 border-transparent mx-auto mb-3" style={{ borderTopColor: '#6366f1' }}></div>
          <p className="text-sm" style={{ color: '#94a3b8' }}>Loading HiHR...</p>
        </div>
      </div>
    );
  }

  // Route based on authentication
  return isAuthenticated ? <DashboardPage /> : <AuthPage />;
};

export default App;
