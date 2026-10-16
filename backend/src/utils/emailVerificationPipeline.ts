/**
 * Email Verification Pipeline
 *
 * Implements the Chain of Responsibility and Strategy Patterns to evaluate
 * email addresses through a sequence of high-performance validation checks.
 */

import {
    DISPOSABLE_DOMAINS,
    FREE_WEBMAIL_DOMAINS,
    ROLE_ACCOUNT_PREFIXES,
    matchHierarchicalDomain,
    matchDisposableMxHost,
    matchDisposableIp,
    matchDisposableNameserver,
    matchDisposableSpf
} from './disposableRegistry';

export type EmailAccountType = 'business' | 'consumer' | 'disposable' | 'invalid';

export interface EmailValidationContext {
    email: string;
    localPart: string;
    domain: string;
    hasMx: boolean;
    exchange?: string;
    mxIps?: string[];
    nameservers?: string[];
    spfRecord?: string;
    dnsError?: string;
}

export interface EmailValidationResult {
    email: string;
    domain: string;
    hasMx: boolean;
    exchange?: string;
    isDisposable: boolean;
    disposableReason?: string;
    accountType: EmailAccountType;
    isRoleAccount: boolean;
    isFreeMail: boolean;
    error?: string;
}

export interface IEmailValidationStrategy {
    name: string;
    execute(context: EmailValidationContext, currentResult: EmailValidationResult): void;
}

/**
 * Strategy 1: Syntax & RFC 5322 Hygiene
 */
export class SyntaxValidationStrategy implements IEmailValidationStrategy {
    name = 'SyntaxValidation';

    private readonly RFC5322_REGEX =
        /^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;

    execute(context: EmailValidationContext, currentResult: EmailValidationResult): void {
        const { email, localPart, domain } = context;

        if (!email || !email.includes('@')) {
            currentResult.error = 'Invalid email syntax (missing @)';
            currentResult.accountType = 'invalid';
            return;
        }

        if (email.length > 254 || localPart.length > 64 || domain.length > 255) {
            currentResult.error = 'Email exceeds maximum allowed character length';
            currentResult.accountType = 'invalid';
            return;
        }

        if (!this.RFC5322_REGEX.test(email)) {
            currentResult.error = 'Invalid RFC 5322 email syntax';
            currentResult.accountType = 'invalid';
            return;
        }

        if (!domain.includes('.') || domain.startsWith('.') || domain.endsWith('.')) {
            currentResult.error = 'Invalid top-level domain syntax';
            currentResult.accountType = 'invalid';
            return;
        }
    }
}

/**
 * Strategy 2: Role-based generic inbox detection
 */
export class RoleAccountStrategy implements IEmailValidationStrategy {
    name = 'RoleAccount';

    execute(context: EmailValidationContext, currentResult: EmailValidationResult): void {
        const cleanLocal = (context.localPart || '').toLowerCase().replace(/[+._-].*$/, '');
        if (ROLE_ACCOUNT_PREFIXES.has(cleanLocal)) {
            currentResult.isRoleAccount = true;
        }
    }
}

/**
 * Strategy 3: O(1) Hierarchical Disposable Domain Matching (120,000+ domains)
 */
export class DisposableDomainStrategy implements IEmailValidationStrategy {
    name = 'DisposableDomain';

    execute(context: EmailValidationContext, currentResult: EmailValidationResult): void {
        const matchedDomain = matchHierarchicalDomain(context.domain, DISPOSABLE_DOMAINS);
        if (matchedDomain) {
            currentResult.isDisposable = true;
            currentResult.disposableReason = `Known disposable provider (${matchedDomain})`;
            currentResult.accountType = 'disposable';
        }
    }
}

/**
 * Strategy 4: Free Consumer Webmail Classification
 */
export class FreeWebmailStrategy implements IEmailValidationStrategy {
    name = 'FreeWebmail';

    execute(context: EmailValidationContext, currentResult: EmailValidationResult): void {
        if (currentResult.isDisposable) return;

        const matchedFreeDomain = matchHierarchicalDomain(context.domain, FREE_WEBMAIL_DOMAINS);
        if (matchedFreeDomain) {
            currentResult.isFreeMail = true;
            currentResult.accountType = 'consumer';
        } else if (currentResult.accountType !== 'invalid') {
            currentResult.accountType = 'business';
        }
    }
}

/**
 * Strategy 5: MX Host Fingerprinting (Catches stealth burner domains)
 */
export class MxFingerprintStrategy implements IEmailValidationStrategy {
    name = 'MxFingerprint';

    execute(context: EmailValidationContext, currentResult: EmailValidationResult): void {
        if (currentResult.isDisposable || !context.exchange) return;

        const fingerprint = matchDisposableMxHost(context.exchange);
        if (fingerprint.isDisposable) {
            currentResult.isDisposable = true;
            currentResult.disposableReason = `MX server points to disposable infrastructure (${fingerprint.provider || context.exchange})`;
            currentResult.accountType = 'disposable';
        }
    }
}

/**
 * Strategy 6: MX Server IP Address Fingerprinting (Catches TempMail rotating fleet)
 */
export class MxIpFingerprintStrategy implements IEmailValidationStrategy {
    name = 'MxIpFingerprint';

    execute(context: EmailValidationContext, currentResult: EmailValidationResult): void {
        if (currentResult.isDisposable || !context.mxIps || context.mxIps.length === 0) return;

        const ipMatch = matchDisposableIp(context.mxIps);
        if (ipMatch.isDisposable) {
            currentResult.isDisposable = true;
            currentResult.disposableReason = `Mail server IP belongs to ${ipMatch.provider || 'disposable cluster'} (${ipMatch.ip})`;
            currentResult.accountType = 'disposable';
        }
    }
}

/**
 * Strategy 7: Nameserver (NS) Fingerprinting (Catches burner DNS fleets like dnsowl.com)
 */
export class NameserverFingerprintStrategy implements IEmailValidationStrategy {
    name = 'NameserverFingerprint';

    execute(context: EmailValidationContext, currentResult: EmailValidationResult): void {
        if (currentResult.isDisposable || !context.nameservers || context.nameservers.length === 0) return;

        const nsMatch = matchDisposableNameserver(context.nameservers);
        if (nsMatch.isDisposable) {
            currentResult.isDisposable = true;
            currentResult.disposableReason = `Domain hosted on burner nameservers (${nsMatch.provider || nsMatch.ns})`;
            currentResult.accountType = 'disposable';
        }
    }
}

/**
 * Strategy 8: Decoy SPF Record Heuristic
 */
export class SpfHeuristicStrategy implements IEmailValidationStrategy {
    name = 'SpfHeuristic';

    execute(context: EmailValidationContext, currentResult: EmailValidationResult): void {
        if (currentResult.isDisposable || !context.spfRecord) return;

        const spfMatch = matchDisposableSpf(context.spfRecord);
        if (spfMatch.isDisposable) {
            currentResult.isDisposable = true;
            currentResult.disposableReason = `Domain matches ${spfMatch.reason || 'decoy SPF signature'}`;
            currentResult.accountType = 'disposable';
        }
    }
}

/**
 * Pipeline Orchestrator (Chain of Responsibility)
 */
export class EmailVerificationPipeline {
    private strategies: IEmailValidationStrategy[] = [];

    constructor() {
        this.registerDefaultStrategies();
    }

    private registerDefaultStrategies(): void {
        this.strategies.push(new SyntaxValidationStrategy());
        this.strategies.push(new RoleAccountStrategy());
        this.strategies.push(new DisposableDomainStrategy());
        this.strategies.push(new FreeWebmailStrategy());
        this.strategies.push(new MxFingerprintStrategy());
        this.strategies.push(new MxIpFingerprintStrategy());
        this.strategies.push(new NameserverFingerprintStrategy());
        this.strategies.push(new SpfHeuristicStrategy());
    }

    /**
     * Executes the verification chain for a given email and DNS context
     */
    evaluate(context: EmailValidationContext): EmailValidationResult {
        const result: EmailValidationResult = {
            email: context.email,
            domain: context.domain,
            hasMx: context.hasMx,
            exchange: context.exchange,
            isDisposable: false,
            accountType: context.hasMx ? 'business' : 'invalid',
            isRoleAccount: false,
            isFreeMail: false,
            error: context.dnsError
        };

        for (const strategy of this.strategies) {
            strategy.execute(context, result);
        }

        // If DNS check explicitly failed and no syntax error occurred
        if (!context.hasMx && !result.error) {
            result.error = context.dnsError || 'No valid MX records found';
            result.accountType = 'invalid';
        }

        return result;
    }
}

export const verificationPipeline = new EmailVerificationPipeline();
