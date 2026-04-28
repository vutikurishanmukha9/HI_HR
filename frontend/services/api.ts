/**
 * Enhanced API Client with retry, timeout, and request interceptors
 */

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:5000/api';

// Configuration
const CONFIG = {
    timeout: 30000, // 30 seconds
    maxRetries: 3,
    retryDelay: 1000, // 1 second base delay
    retryableStatuses: [408, 429, 500, 502, 503, 504],
};

// Types
interface RequestConfig extends RequestInit {
    timeout?: number;
    retries?: number;
    skipRetry?: boolean;
}

interface ApiError extends Error {
    status?: number;
    code?: string;
}

/**
 * Create an AbortController with timeout
 */
function createTimeoutController(timeout: number): { controller: AbortController; timeoutId: NodeJS.Timeout } {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeout);
    return { controller, timeoutId };
}

/**
 * Delay helper for retry
 */
function delay(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Enhanced API Client
 */
class ApiClient {
    private baseURL: string;
    private token: string | null = null;

    constructor(baseURL: string) {
        this.baseURL = baseURL;
        this.token = localStorage.getItem('accessToken');
    }

    setToken(token: string | null) {
        this.token = token;
        if (token) {
            localStorage.setItem('accessToken', token);
        } else {
            localStorage.removeItem('accessToken');
        }
    }

    /**
     * Core request method with retry and timeout support
     */
    private async request<T>(
        endpoint: string,
        options: RequestConfig = {}
    ): Promise<T> {
        const {
            timeout = CONFIG.timeout,
            retries = CONFIG.maxRetries,
            skipRetry = false,
            ...fetchOptions
        } = options;

        const headers: HeadersInit = {
            'Content-Type': 'application/json',
            ...fetchOptions.headers,
        };

        if (this.token) {
            (headers as Record<string, string>)['Authorization'] = `Bearer ${this.token}`;
        }

        let lastError: ApiError | null = null;

        for (let attempt = 0; attempt <= retries; attempt++) {
            const { controller, timeoutId } = createTimeoutController(timeout);

            try {
                const response = await fetch(`${this.baseURL}${endpoint}`, {
                    ...fetchOptions,
                    headers,
                    signal: controller.signal,
                });

                clearTimeout(timeoutId);

                // Parse response
                let data: any;
                const contentType = response.headers.get('content-type');

                if (contentType?.includes('application/json')) {
                    try {
                        data = await response.json();
                    } catch {
                        data = { error: 'Invalid JSON response from server' };
                    }
                } else {
                    const text = await response.text();
                    data = { error: text || 'Unexpected server response' };
                }

                if (!response.ok) {
                    const error: ApiError = new Error(
                        this.extractErrorMessage(data, response)
                    );
                    error.status = response.status;
                    error.code = data?.error?.code;

                    // Check if should retry
                    if (
                        !skipRetry &&
                        attempt < retries &&
                        CONFIG.retryableStatuses.includes(response.status)
                    ) {
                        lastError = error;
                        await delay(CONFIG.retryDelay * Math.pow(2, attempt)); // Exponential backoff
                        continue;
                    }

                    throw error;
                }

                return data;
            } catch (error: any) {
                clearTimeout(timeoutId);

                if (error.name === 'AbortError') {
                    const timeoutError: ApiError = new Error('Request timeout');
                    timeoutError.code = 'TIMEOUT';

                    if (!skipRetry && attempt < retries) {
                        lastError = timeoutError;
                        await delay(CONFIG.retryDelay * Math.pow(2, attempt));
                        continue;
                    }
                    throw timeoutError;
                }

                // Network errors - retry
                if (!skipRetry && attempt < retries && error.message.includes('fetch')) {
                    lastError = error;
                    await delay(CONFIG.retryDelay * Math.pow(2, attempt));
                    continue;
                }

                throw error;
            }
        }

        throw lastError || new Error('Request failed after retries');
    }

    /**
     * Extract error message from various response formats
     */
    private extractErrorMessage(data: any, response: Response): string {
        if (data?.error?.message) return data.error.message;
        if (typeof data?.error === 'string') return data.error;
        if (data?.message) return data.message;
        return `Server error: ${response.status} ${response.statusText}`;
    }

    // ============= Auth Endpoints =============

    async register(email: string, password: string, fullName: string) {
        const data = await this.request<{
            user: any;
            accessToken: string;
            refreshToken: string;
        }>('/auth/register', {
            method: 'POST',
            body: JSON.stringify({ email, password, fullName }),
            skipRetry: true, // Don't retry auth requests
        });
        this.setToken(data.accessToken);
        localStorage.setItem('refreshToken', data.refreshToken);
        return data;
    }

    async login(email: string, password: string) {
        const data = await this.request<{
            user: any;
            accessToken: string;
            refreshToken: string;
        }>('/auth/login', {
            method: 'POST',
            body: JSON.stringify({ email, password }),
            skipRetry: true,
        });
        this.setToken(data.accessToken);
        localStorage.setItem('refreshToken', data.refreshToken);
        return data;
    }

    async getMe() {
        return this.request<{ user: any }>('/auth/me');
    }

    async refreshToken() {
        const refreshToken = localStorage.getItem('refreshToken');
        if (!refreshToken) {
            throw new Error('No refresh token available');
        }

        const data = await this.request<{ accessToken: string }>('/auth/refresh', {
            method: 'POST',
            body: JSON.stringify({ refreshToken }),
            skipRetry: true,
        });

        this.setToken(data.accessToken);
        return data;
    }

    async logout() {
        try {
            await this.request('/auth/logout', { method: 'POST', skipRetry: true });
        } catch {
            // Ignore errors - still clear local token
        } finally {
            this.setToken(null);
            localStorage.removeItem('refreshToken');
        }
    }

    // ============= Credential Endpoints =============

    async saveCredential(email: string, appPassword: string, isDefault?: boolean) {
        return this.request<{ credential: any }>('/credentials', {
            method: 'POST',
            body: JSON.stringify({ email, appPassword, isDefault }),
        });
    }

    async getCredentials() {
        return this.request<{ credentials: any[] }>('/credentials');
    }

    async deleteCredential(id: string) {
        return this.request<{ message: string }>(`/credentials/${id}`, {
            method: 'DELETE',
        });
    }

    // ============= Campaign Endpoints =============

    async createCampaign(data: {
        name: string;
        subject: string;
        body: string;
        recipients: any[];
        scheduledTime?: string;
        batchSize?: number;
        batchDelay?: number;
    }) {
        return this.request<{ campaign: any }>('/campaigns', {
            method: 'POST',
            body: JSON.stringify(data),
        });
    }

    async getCampaigns(params?: { page?: number; limit?: number; status?: string }) {
        const query = new URLSearchParams(params as any).toString();
        return this.request<{
            campaigns: any[];
            total: number;
            page: number;
            totalPages: number;
        }>(`/campaigns?${query}`);
    }

    async getCampaign(id: string) {
        return this.request<{ campaign: any }>(`/campaigns/${id}`);
    }

    async updateCampaign(id: string, data: any) {
        return this.request<{ campaign: any }>(`/campaigns/${id}`, {
            method: 'PATCH',
            body: JSON.stringify(data),
        });
    }

    async deleteCampaign(id: string) {
        return this.request<{ message: string }>(`/campaigns/${id}`, {
            method: 'DELETE',
        });
    }

    async uploadRecipients(file: File) {
        const formData = new FormData();
        formData.append('file', file);

        const { controller, timeoutId } = createTimeoutController(60000); // 60s for uploads

        try {
            const response = await fetch(`${this.baseURL}/campaigns/upload-recipients`, {
                method: 'POST',
                headers: this.token ? { Authorization: `Bearer ${this.token}` } : {},
                body: formData,
                signal: controller.signal,
            });

            clearTimeout(timeoutId);
            const data = await response.json();

            if (!response.ok) {
                throw new Error(data.error || 'Upload failed');
            }

            return data;
        } catch (error: any) {
            clearTimeout(timeoutId);
            if (error.name === 'AbortError') {
                throw new Error('Upload timeout - file may be too large');
            }
            throw error;
        }
    }

    /**
     * Send campaign emails via backend (secure - no credentials exposed)
     */
    async sendCampaign(data: {
        credentialEmail: string;
        subject: string;
        body: string;
        recipients: Array<{
            email: string;
            fullName: string;
            companyName: string;
            jobTitle?: string;
        }>;
        attachments?: File[];
        batchSize?: number;
        batchDelay?: number;
    }) {
        // Convert File objects to base64 for JSON transport
        let attachmentsData: Array<{ filename: string; content: string; contentType: string }> | undefined;

        if (data.attachments && data.attachments.length > 0) {
            attachmentsData = await Promise.all(
                data.attachments.map(async (file) => {
                    const buffer = await file.arrayBuffer();
                    const base64 = btoa(
                        new Uint8Array(buffer).reduce((d, byte) => d + String.fromCharCode(byte), '')
                    );
                    return {
                        filename: file.name,
                        content: base64,
                        contentType: file.type || 'application/octet-stream',
                    };
                })
            );
        }

        return this.request<{
            message: string;
            results: Array<{ email: string; status: 'sent' | 'failed'; error?: string }>;
            summary: { total: number; sent: number; failed: number };
        }>('/campaigns/send', {
            method: 'POST',
            body: JSON.stringify({
                ...data,
                attachments: attachmentsData,
            }),
            timeout: 300000, // 5 minutes for large campaigns
            skipRetry: true, // Don't retry campaign sends
        });
    }

    // ============= Generic HTTP Methods =============

    async get<T>(endpoint: string): Promise<T> {
        return this.request<T>(endpoint, { method: 'GET' });
    }

    async post<T>(endpoint: string, data?: any): Promise<T> {
        return this.request<T>(endpoint, {
            method: 'POST',
            body: data ? JSON.stringify(data) : undefined,
        });
    }
}

export const apiClient = new ApiClient(API_BASE_URL);
export default apiClient;
