import React from 'react';
import { AuthProvider, useAuth } from './contexts/AuthContext';
import { ToastProvider } from './components/Toast';
import AuthPage from './pages/AuthPage';
import DashboardPage from './pages/DashboardPage';

/**
 * App Content Component
 * Handles routing between Auth and Dashboard based on authentication state
 */
const AppContent: React.FC = () => {
  const { isAuthenticated, isLoading } = useAuth();

  // Show loading screen while checking auth
  if (isLoading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-canvas">
        <div className="text-center">
          <div className="inline-flex items-center justify-center w-11 h-11 rounded-xl bg-accent text-white font-bold text-lg mb-4 shadow-sm">
            Hi
          </div>
          <div className="spinner mx-auto mb-3" />
          <p className="text-sm text-ink-muted">Loading workspace…</p>
        </div>
      </div>
    );
  }

  // Route based on authentication
  return isAuthenticated ? <DashboardPage /> : <AuthPage />;
};

/**
 * Main App Component
 * Wraps top-level context providers (ToastProvider, AuthProvider)
 * to guarantee robust provider inheritance during Vite HMR updates
 */
const App: React.FC = () => {
  return (
    <ToastProvider>
      <AuthProvider>
        <AppContent />
      </AuthProvider>
    </ToastProvider>
  );
};

export default App;
