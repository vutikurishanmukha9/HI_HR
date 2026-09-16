import dns from 'dns';
const dnsPromises = dns.promises;

export interface DomainCheckResult {
    domain: string;
    hasMx: boolean;
    exchange?: string;
    error?: string;
}

export class EmailValidatorService {
    private cache: Map<string, { result: DomainCheckResult; cachedAt: number }> = new Map();
    private readonly CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour

    /**
     * Checks if a single domain has valid MX records
     */
    async checkDomainMx(domain: string): Promise<DomainCheckResult> {
        const cleanDomain = (domain || '').trim().toLowerCase();
        if (!cleanDomain || !cleanDomain.includes('.')) {
            return { domain: cleanDomain, hasMx: false, error: 'Invalid domain syntax' };
        }

        // Check cache
        const cached = this.cache.get(cleanDomain);
        if (cached && (Date.now() - cached.cachedAt) < this.CACHE_TTL_MS) {
            return cached.result;
        }

        try {
            const mxRecords = await dnsPromises.resolveMx(cleanDomain);
            if (mxRecords && mxRecords.length > 0) {
                // Sort by priority (lowest number = highest priority)
                mxRecords.sort((a, b) => a.priority - b.priority);
                const result: DomainCheckResult = {
                    domain: cleanDomain,
                    hasMx: true,
                    exchange: mxRecords[0].exchange,
                };
                this.cache.set(cleanDomain, { result, cachedAt: Date.now() });
                return result;
            } else {
                const result: DomainCheckResult = {
                    domain: cleanDomain,
                    hasMx: false,
                    error: 'No MX records found for domain',
                };
                this.cache.set(cleanDomain, { result, cachedAt: Date.now() });
                return result;
            }
        } catch (error: any) {
            let errorMsg = 'DNS lookup failed';
            if (error.code === 'ENOTFOUND' || error.code === 'NXDOMAIN') {
                errorMsg = 'Domain does not exist (NXDOMAIN)';
            } else if (error.code === 'ENODATA') {
                errorMsg = 'Domain exists but has no mail server configured';
            } else if (error.code === 'ETIMEOUT') {
                errorMsg = 'DNS lookup timed out';
            }

            const result: DomainCheckResult = {
                domain: cleanDomain,
                hasMx: false,
                error: errorMsg,
            };
            this.cache.set(cleanDomain, { result, cachedAt: Date.now() });
            return result;
        }
    }

    /**
     * Concurrently checks MX records for multiple email addresses / domains
     */
    async validateEmailsMx(emails: string[]): Promise<Record<string, DomainCheckResult>> {
        const uniqueDomains = new Set<string>();
        const emailToDomainMap: Record<string, string> = {};

        for (const email of emails) {
            if (email && email.includes('@')) {
                const domain = email.split('@')[1].trim().toLowerCase();
                uniqueDomains.add(domain);
                emailToDomainMap[email] = domain;
            }
        }

        const domainResults: Record<string, DomainCheckResult> = {};
        const domainList = Array.from(uniqueDomains);

        // Process in batches of 10 concurrent lookups
        const BATCH_SIZE = 10;
        for (let i = 0; i < domainList.length; i += BATCH_SIZE) {
            const batch = domainList.slice(i, i + BATCH_SIZE);
            const batchPromises = batch.map(d => this.checkDomainMx(d));
            const batchResults = await Promise.all(batchPromises);
            batchResults.forEach(r => {
                domainResults[r.domain] = r;
            });
        }

        // Map back to individual emails
        const emailResults: Record<string, DomainCheckResult> = {};
        for (const email of emails) {
            const domain = emailToDomainMap[email];
            if (domain && domainResults[domain]) {
                emailResults[email] = domainResults[domain];
            } else {
                emailResults[email] = {
                    domain: domain || 'unknown',
                    hasMx: false,
                    error: 'Invalid email address format',
                };
            }
        }

        return emailResults;
    }
}

export default new EmailValidatorService();
