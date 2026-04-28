import React, { createContext, useContext, useState, useEffect, ReactNode } from 'react';
import apiClient from '../services/api';

interface User {
    id: string;
    email: string;
    fullName: string;
}

interface AuthContextType {
    user: User | null;
    isAuthenticated: boolean;
    isLoading: boolean;
    login: (email: string, password: string) => Promise<void>;
    register: (email: string, password: string, fullName: string) => Promise<void>;
    logout: () => void;
}

const AuthContext = createContext<AuthContextType | undefined>(undefined);

export const useAuth = () => {
    const context = useContext(AuthContext);
    if (!context) {
        throw new Error('useAuth must be used within AuthProvider');
    }
    return context;
};

interface AuthProviderProps {
    children: ReactNode;
}

export const AuthProvider: React.FC<AuthProviderProps> = ({ children }) => {
    const [user, setUser] = useState<User | null>(null);
    const [isLoading, setIsLoading] = useState(true);

    useEffect(() => {
        // Check if user is logged in on mount
        const checkAuth = async () => {
            try {
                const token = localStorage.getItem('accessToken');
                if (token) {
                    // Only call API if token exists
                    const { user } = await apiClient.getMe();
                    setUser(user);
                }
            } catch (error) {
                // Token invalid or expired, clear it
                console.log('Auth check failed:', error);
                localStorage.removeItem('accessToken');
                apiClient.setToken(null);
            } finally {
                setIsLoading(false);
            }
        };

        checkAuth();
    }, []);

    const login = async (email: string, password: string) => {
        const { user } = await apiClient.login(email, password);
        setUser(user);
    };

    const register = async (email: string, password: string, fullName: string) => {
        const { user } = await apiClient.register(email, password, fullName);
        setUser(user);
    };

    const logout = () => {
        apiClient.logout();
        setUser(null);
    };

    return (
        <AuthContext.Provider
            value={{
                user,
                isAuthenticated: !!user,
                isLoading,
                login,
                register,
                logout,
            }}
        >
            {children}
        </AuthContext.Provider>
    );
};
