/**
 * Enterprise Disposable & Temporary Email Registry
 *
 * Implements high-performance O(1) in-memory lookup sets and host signature
 * matchers to detect throwaway inboxes, disposable services, free consumer webmail,
 * and generic role accounts.
 */

// Load comprehensive industry-standard dataset of 120,000+ disposable domains
let packageDomains: string[] = [];
try {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    packageDomains = require('disposable-email-domains');
} catch (err) {
    packageDomains = [];
}

let packageWildcards: string[] = [];
try {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    packageWildcards = require('disposable-email-domains/wildcard.json');
} catch (err) {
    packageWildcards = [];
}

// Curated supplementary aliases, international temp services, and TempMail fleet domains
const SUPPLEMENTARY_DISPOSABLE_DOMAINS: string[] = [
    // TempMail rotating burner fleet (e.g. hiredify.com, omanarts.com, kingdais.com)
    'hiredify.com', 'omanarts.com', 'kingdais.com', 'zoomugo.com', 'ncleap.com', 'shopdevo.com',
    'hi2.in', 'mailinator.com', 'mailinater.com', 'suremail.info', 'spamherelots.com', 'binkmail.com',
    'safetymail.info', 'chammy.info', 'devnullmail.com', 'letthemeatspam.com', 'mailin8r.com',
    'mailinator2.com', 'notmailinator.com', 'reallymymail.com', 'reconmail.com', 'thisisnotmyrealemail.com',
    'tradermail.info', 'veryrealemail.com', 'zippymail.info', 'mailinator.net', 'mailinator.org',
    'guerrillamail.com', 'guerrillamailblock.com', 'guerrillamail.net', 'guerrillamail.biz',
    'guerrillamail.org', 'guerrillamail.de', 'sharklasers.com', 'grr.la', 'pokemail.net', 'spam4.me',
    '10minutemail.com', '10minutemail.net', '10minutemail.org', '10minutemail.co.uk',
    '10minemail.com', '10minutemailbox.com', '10minutesmail.com', '10minutemail.be',
    'temp-mail.org', 'tempmail.com', 'temp-mail.ru', 'tempmail.net', 'tempmailaddress.com',
    'tempail.com', 'tempinbox.com', 'tempemail.co', 'temp-mail.io', 'tempmailo.com',
    'tempmailgen.com', 'tmpmail.org', 'tmpmail.net',
    'yopmail.com', 'yopmail.fr', 'yopmail.net', 'cool.fr.nf', 'courriel.fr.nf',
    'jetable.fr.nf', 'nospam.ze.tc', 'nomail.xl.cx', 'mega.zik.dj', 'speed.1s.fr',
    'hide.biz.st', 'mymail.infos.st',
    'trashmail.com', 'trashmail.net', 'trashmail.me', 'trashmail.org', 'trashmail.at',
    'trashmail.io', 'trash-mail.com', 'trash-mail.at', 'trashmail.club',
    'throwawaymail.com', 'throwaway.email', 'dispostable.com', 'burnermail.io',
    'burnmail.at', 'getairmail.com', 'airmail.news', 'fakeinbox.com', 'fakemailgenerator.com',
    'mohmal.com', 'mohmal.im', 'crazymailing.com', 'crazymail.com', 'maildrop.cc',
    'nada.ltd', 'getnada.com', 'inboxkitten.com', 'inboxbear.com', 'mailpoof.com',
    'byom.de', 'mailnesia.com', 'mailcatch.com', 'mintemail.com', 'spambox.us',
    'meltmail.com', 'deadaddress.com', 'mytemp.email', 'emailondeck.com', 'luxusmail.org',
    'generator.email', 'anonbox.net', 'anonymbox.com', 'antichef.com', 'baxomale.ht.cx',
    'discard.email', 'discardmail.com', 'disposablemail.com', 'disposableemailaddresses.com'
];

// Unified fast O(1) hash set with 120,000+ domains
const unifiedDisposableSet = new Set<string>(packageDomains);
SUPPLEMENTARY_DISPOSABLE_DOMAINS.forEach(d => unifiedDisposableSet.add(d.toLowerCase().trim()));

export const DISPOSABLE_DOMAINS: ReadonlySet<string> = unifiedDisposableSet;
export const DISPOSABLE_WILDCARDS: ReadonlySet<string> = new Set<string>(packageWildcards.map(w => w.toLowerCase().trim()));

// Known disposable mail server IP addresses (TempMail, GuerrillaMail, etc.)
export const DISPOSABLE_MX_IPS: ReadonlySet<string> = new Set<string>([
    '134.199.178.234', // Temp-Mail.org / TempMail inbound cluster (hiredify.com, omanarts.com, kingdais.com, etc.)
]);

// Known disposable Nameserver patterns (catches stealth burner domain fleets)
export const DISPOSABLE_NAMESERVER_SIGNATURES: readonly { pattern: RegExp; provider: string }[] = [
    { pattern: /dnsowl\.com$/i, provider: 'Temp-Mail.org (dnsowl)' },
    { pattern: /temp-mail/i, provider: 'TempMail' },
    { pattern: /mailinator/i, provider: 'Mailinator' },
    { pattern: /guerrillamail/i, provider: 'GuerrillaMail' },
    { pattern: /yopmail/i, provider: 'Yopmail' },
    { pattern: /trashmail/i, provider: 'Trashmail' },
];

// Known disposable MX host signatures (catches stealth burner domains with custom hostnames)
export const DISPOSABLE_MX_SIGNATURES: readonly { pattern: RegExp; provider: string }[] = [
    { pattern: /mailinator\.com$/i, provider: 'Mailinator' },
    { pattern: /yopmail\.(com|net|fr)$/i, provider: 'Yopmail' },
    { pattern: /guerrillamail(block)?\.(com|net|biz|org)$/i, provider: 'GuerrillaMail' },
    { pattern: /sharklasers\.com$/i, provider: 'GuerrillaMail / Sharklasers' },
    { pattern: /trashmail\.(com|net|me|at|org|club)$/i, provider: 'Trashmail' },
    { pattern: /temp-mail\.(org|io|ru|net)$/i, provider: 'TempMail' },
    { pattern: /10minutemail\.(com|net|org)$/i, provider: '10MinuteMail' },
    { pattern: /dispostable\.com$/i, provider: 'Dispostable' },
    { pattern: /burnermail\.io$/i, provider: 'BurnerMail' },
    { pattern: /mohmal\.(com|im|in|tech)$/i, provider: 'Mohmal' },
    { pattern: /fakeinbox\.com$/i, provider: 'FakeInbox' },
    { pattern: /maildrop\.cc$/i, provider: 'MailDrop' },
    { pattern: /mailsac\.com$/i, provider: 'Mailsac' },
    { pattern: /mailslurp\.(com|io)$/i, provider: 'MailSlurp' },
    { pattern: /emailondeck\.com$/i, provider: 'EmailOnDeck' },
    { pattern: /throwawaymail\.com$/i, provider: 'ThrowAwayMail' },
    { pattern: /getairmail\.com$/i, provider: 'GetAirMail' },
    { pattern: /inboxkitten\.com$/i, provider: 'InboxKitten' },
    { pattern: /hi2\.in$/i, provider: 'FakeMail (hi2.in)' },
];

// Major free consumer webmail domains
export const FREE_WEBMAIL_DOMAINS: ReadonlySet<string> = new Set<string>([
    'gmail.com', 'googlemail.com',
    'yahoo.com', 'yahoo.co.uk', 'yahoo.fr', 'yahoo.de', 'yahoo.es', 'yahoo.in', 'yahoo.co.jp', 'ymail.com', 'rocketmail.com',
    'outlook.com', 'hotmail.com', 'live.com', 'msn.com', 'passport.com', 'hotmail.co.uk', 'hotmail.fr', 'live.co.uk',
    'icloud.com', 'me.com', 'mac.com',
    'proton.me', 'protonmail.com', 'pm.me',
    'zoho.com', 'zohomail.com',
    'aol.com', 'aim.com',
    'gmx.com', 'gmx.net', 'gmx.de',
    'mail.com', 'email.com',
    'mail.ru', 'inbox.ru', 'list.ru', 'bk.ru',
    'yandex.com', 'yandex.ru', 'ya.ru',
    'tutanota.com', 'tutamail.com', 'tuta.io',
    'fastmail.com', 'fastmail.fm',
    'web.de', 't-online.de',
    'orange.fr', 'wanadoo.fr', 'free.fr', 'sfr.fr',
    'libero.it', 'virgilio.it',
    'rediffmail.com'
]);

// Known generic role-based address prefixes (sales, info, support, hr, etc.)
export const ROLE_ACCOUNT_PREFIXES: ReadonlySet<string> = new Set<string>([
    'admin', 'administrator', 'info', 'information', 'support', 'help', 'helpdesk',
    'sales', 'contact', 'contactus', 'office', 'team', 'jobs', 'careers', 'recruiting',
    'recruitment', 'hr', 'humanresources', 'billing', 'accounts', 'accounting', 'finance',
    'marketing', 'press', 'media', 'pr', 'legal', 'compliance', 'privacy', 'security',
    'no-reply', 'noreply', 'donotreply', 'webmaster', 'postmaster', 'hostmaster', 'abuse',
    'operations', 'devops', 'engineering', 'general', 'service', 'services', 'feedback'
]);

/**
 * Traverses domain subdomains hierarchically (TLD+1, TLD+2)
 * to catch domains like sub.mailinator.com or anything.hi2.in
 */
export function matchHierarchicalDomain(domain: string, set: ReadonlySet<string>): string | null {
    const clean = (domain || '').toLowerCase().trim();
    if (!clean) return null;

    if (set.has(clean)) {
        return clean;
    }

    // Check wildcard match
    if (DISPOSABLE_WILDCARDS.has(clean)) {
        return clean;
    }

    const parts = clean.split('.');
    for (let i = 1; i < parts.length - 1; i++) {
        const candidate = parts.slice(i).join('.');
        if (set.has(candidate) || DISPOSABLE_WILDCARDS.has(candidate)) {
            return candidate;
        }
    }

    return null;
}

/**
 * Inspects a mail exchanger hostname for known disposable signatures
 */
export function matchDisposableMxHost(exchange: string): { isDisposable: boolean; provider?: string } {
    if (!exchange) return { isDisposable: false };
    const cleanExchange = exchange.toLowerCase().trim().replace(/\.$/, '');

    for (const item of DISPOSABLE_MX_SIGNATURES) {
        if (item.pattern.test(cleanExchange)) {
            return {
                isDisposable: true,
                provider: item.provider
            };
        }
    }

    return { isDisposable: false };
}

/**
 * Inspects resolved MX IP addresses against known disposable cluster IPs
 */
export function matchDisposableIp(ips: readonly string[]): { isDisposable: boolean; ip?: string; provider?: string } {
    for (const ip of ips) {
        if (DISPOSABLE_MX_IPS.has(ip)) {
            return {
                isDisposable: true,
                ip,
                provider: 'TempMail Infrastructure'
            };
        }
    }
    return { isDisposable: false };
}

/**
 * Inspects domain nameservers for known burner domain DNS providers
 */
export function matchDisposableNameserver(nameservers: readonly string[]): { isDisposable: boolean; ns?: string; provider?: string } {
    for (const ns of nameservers) {
        const cleanNs = (ns || '').toLowerCase().trim().replace(/\.$/, '');
        for (const item of DISPOSABLE_NAMESERVER_SIGNATURES) {
            if (item.pattern.test(cleanNs)) {
                return {
                    isDisposable: true,
                    ns: cleanNs,
                    provider: item.provider
                };
            }
        }
    }
    return { isDisposable: false };
}

/**
 * Inspects SPF records for decoy/fake disposable patterns (e.g. v=spf1 ip4:1.1.1.1 -all)
 */
export function matchDisposableSpf(spf: string): { isDisposable: boolean; reason?: string } {
    if (!spf) return { isDisposable: false };
    const cleanSpf = spf.toLowerCase().trim();
    if (cleanSpf.includes('ip4:1.1.1.1') || cleanSpf.includes('ip4:0.0.0.0')) {
        return {
            isDisposable: true,
            reason: 'Decoy SPF record detected'
        };
    }
    return { isDisposable: false };
}
