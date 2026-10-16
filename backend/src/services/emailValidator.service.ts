import dns from 'dns';
import {
    verificationPipeline,
    EmailAccountType
} from '../utils/emailVerificationPipeline';

const dnsPromises = dns.promises;

export interface DomainCheckResult {
    domain: string;
    hasMx: boolean;
    exchange?: string;
    error?: string;
    isDisposable: boolean;
    disposableReason?: string;
    accountType: EmailAccountType;
    isRoleAccount?: boolean;
    isFreeMail?: boolean;
    email?: string;
}

interface RawDnsResult {
    domain: string;
    hasMx: boolean;
    exchange?: string;
    mxIps?: string[];
    nameservers?: string[];
    spfRecord?: string;
    error?: string;
}

export class EmailValidatorService {
    private cache: Map<string, { result: RawDnsResult; cachedAt: number }> = new Map();
    private readonly CACHE_TTL_MS = 60 * 60 * 1000; // 1 hour

    /**
     * Resolves DNS MX, A, NS, and TXT records with caching and error handling
     */
    async resolveDomainDns(domain: string): Promise<RawDnsResult> {
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
                const primaryExchange = mxRecords[0].exchange;

                // Concurrently query A record for mail server, NS records, and SPF TXT
                const [mxIpsResult, nsResult, txtResult] = await Promise.allSettled([
                    dnsPromises.resolve4(primaryExchange),
                    dnsPromises.resolveNs(cleanDomain),
                    dnsPromises.resolveTxt(cleanDomain)
                ]);

                const mxIps = mxIpsResult.status === 'fulfilled' ? mxIpsResult.value : [];
                const nameservers = nsResult.status === 'fulfilled' ? nsResult.value : [];
                let spfRecord: string | undefined;

                if (txtResult.status === 'fulfilled' && Array.isArray(txtResult.value)) {
                    const flat = txtResult.value.flat();
                    spfRecord = flat.find(t => typeof t === 'string' && t.startsWith('v=spf1'));
                }

                const result: RawDnsResult = {
                    domain: cleanDomain,
                    hasMx: true,
                    exchange: primaryExchange,
                    mxIps,
                    nameservers,
                    spfRecord,
                };
                this.cache.set(cleanDomain, { result, cachedAt: Date.now() });
                return result;
            } else {
                const result: RawDnsResult = {
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

            const result: RawDnsResult = {
                domain: cleanDomain,
                hasMx: false,
                error: errorMsg,
            };
            this.cache.set(cleanDomain, { result, cachedAt: Date.now() });
            return result;
        }
    }

    /**
     * Evaluates a single email or domain through the verification pipeline
     */
    async checkDomainMx(domain: string): Promise<DomainCheckResult> {
        const clean = (domain || '').trim().toLowerCase();
        const dnsResult = await this.resolveDomainDns(clean);

        const evaluated = verificationPipeline.evaluate({
            email: `verify@${clean}`,
            localPart: 'verify',
            domain: clean,
            hasMx: dnsResult.hasMx,
            exchange: dnsResult.exchange,
            mxIps: dnsResult.mxIps,
            nameservers: dnsResult.nameservers,
            spfRecord: dnsResult.spfRecord,
            dnsError: dnsResult.error
        });

        return {
            domain: evaluated.domain,
            hasMx: evaluated.hasMx,
            exchange: evaluated.exchange,
            error: evaluated.error,
            isDisposable: evaluated.isDisposable,
            disposableReason: evaluated.disposableReason,
            accountType: evaluated.accountType,
            isFreeMail: evaluated.isFreeMail
        };
    }

    /**
     * Concurrently checks MX records and evaluates disposable/role status for emails
     */
    async validateEmailsMx(emails: string[]): Promise<Record<string, DomainCheckResult>> {
        const uniqueDomains = new Set<string>();
        const parsedEmails: { original: string; localPart: string; domain: string }[] = [];

        for (const email of emails) {
            const clean = (email || '').trim().toLowerCase();
            if (clean && clean.includes('@')) {
                const atIdx = clean.indexOf('@');
                const localPart = clean.slice(0, atIdx);
                const domain = clean.slice(atIdx + 1);
                uniqueDomains.add(domain);
                parsedEmails.push({ original: email, localPart, domain });
            } else {
                parsedEmails.push({ original: email, localPart: '', domain: '' });
            }
        }

        // Concurrent DNS resolution in batches of 10
        const domainList = Array.from(uniqueDomains);
        const dnsMap: Record<string, RawDnsResult> = {};
        const BATCH_SIZE = 10;

        for (let i = 0; i < domainList.length; i += BATCH_SIZE) {
            const batch = domainList.slice(i, i + BATCH_SIZE);
            const batchResults = await Promise.all(batch.map(d => this.resolveDomainDns(d)));
            batchResults.forEach(r => {
                dnsMap[r.domain] = r;
            });
        }

        // Evaluate each email through the verification pipeline
        const emailResults: Record<string, DomainCheckResult> = {};

        for (const item of parsedEmails) {
            if (!item.domain || !dnsMap[item.domain]) {
                emailResults[item.original] = {
                    email: item.original,
                    domain: item.domain || 'unknown',
                    hasMx: false,
                    isDisposable: false,
                    accountType: 'invalid',
                    error: 'Invalid email address format',
                };
                continue;
            }

            const dnsRes = dnsMap[item.domain];
            const evaluated = verificationPipeline.evaluate({
                email: item.original,
                localPart: item.localPart,
                domain: item.domain,
                hasMx: dnsRes.hasMx,
                exchange: dnsRes.exchange,
                mxIps: dnsRes.mxIps,
                nameservers: dnsRes.nameservers,
                spfRecord: dnsRes.spfRecord,
                dnsError: dnsRes.error
            });

            emailResults[item.original] = {
                email: item.original,
                domain: evaluated.domain,
                hasMx: evaluated.hasMx,
                exchange: evaluated.exchange,
                error: evaluated.error,
                isDisposable: evaluated.isDisposable,
                disposableReason: evaluated.disposableReason,
                accountType: evaluated.accountType,
                isRoleAccount: evaluated.isRoleAccount,
                isFreeMail: evaluated.isFreeMail
            };
        }

        return emailResults;
    }
}

export default new EmailValidatorService();
