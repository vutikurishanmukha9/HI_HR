import { SpamReport, SpamIssue, DetectedSpamWord, DeliverabilityMetrics } from '../types';

interface TriggerWordDef {
    phrase: string;
    category: 'Scam & Fraud' | 'Urgency & Pressure' | 'Deceptive Phrasing' | 'Deliverability Flag';
    severity: 'high' | 'medium';
    suggestion: string;
    matcher: RegExp;
    isBenign?: (fullContext: string) => boolean;
}

/**
 * Technical and business acronyms exempt from all-caps shouting penalties
 */
const ALLOWED_ACRONYMS = new Set([
    'AWS', 'GCP', 'AZURE', 'API', 'REST', 'HTML', 'CSS', 'JSON', 'SQL', 'HTTP', 'HTTPS',
    'SAAS', 'B2B', 'B2C', 'CRM', 'ERP', 'SDK', 'CEO', 'CTO', 'CFO', 'COO', 'CPO', 'VP',
    'HR', 'PR', 'UI', 'UX', 'AI', 'ML', 'LLM', 'ROI', 'KPI', 'KPIS', 'SLA', 'OKR', 'OKRS',
    'GDPR', 'HIPAA', 'SOC2', 'CI/CD', 'CICD', 'QA', 'SRE', 'SEO', 'SEM', 'CTR'
]);

/**
 * Generic URL shorteners known to harm email deliverability
 */
const SUSPICIOUS_SHORTENER_REGEX = /\b(bit\.ly|tinyurl\.com|goo\.gl|ow\.ly|is\.gd|buff\.ly|t\.co|rebrand\.ly|bl\.ink|cutt\.ly)\b/i;

/**
 * Intelligent trigger word rules with context awareness to eliminate false positives
 */
const TRIGGER_DEFINITIONS: TriggerWordDef[] = [
    // --- 1. Scam & Fraud (High Risk) ---
    {
        phrase: '100% free',
        category: 'Scam & Fraud',
        severity: 'high',
        suggestion: "Use 'complimentary' or 'included at no cost'",
        matcher: /\b(100%\s*free|free\s*money|free\s*gift|free\s*cash|free\s*membership)\b/i,
    },
    {
        phrase: 'congratulations (lottery/prize scam)',
        category: 'Scam & Fraud',
        severity: 'high',
        suggestion: "Avoid lottery or prize claims; congratulating professional milestones like 'congratulations on the launch' is safe",
        matcher: /\bcongratulations\b/i,
        isBenign: (context: string) => {
            // Check for scam phrases
            const scamPatterns = [
                /congratulations.*?\b(won|winner|lottery|claim\s+your|cash\s+prize|selected\s+for\s+\$)\b/i
            ];
            if (scamPatterns.some(p => p.test(context))) {
                return false;
            }
            // Whitelist legitimate professional praise
            const benignPatterns = [
                /congratulations\s+on/i,
                /congratulations\s+to/i,
                /congratulations\s+team/i,
                /congratulations\s+\{[a-z0-9_.-]+\}/i,
                /congratulations[!.,\s]/i,
                /congrats\s+on/i,
                /congrats\s+to/i,
            ];
            return benignPatterns.some(p => p.test(context));
        }
    },
    {
        phrase: 'pure profit',
        category: 'Scam & Fraud',
        severity: 'high',
        suggestion: "Use 'operating margin' or 'net revenue'",
        matcher: /\b(pure\s+profit|guaranteed\s+profit|instant\s+profit)\b/i,
    },
    {
        phrase: 'make money fast',
        category: 'Scam & Fraud',
        severity: 'high',
        suggestion: "Use business metrics like 'drive revenue' or 'accelerate pipeline'",
        matcher: /\b(make\s+money(\s+fast)?|earn\s+(extra\s+)?cash|fast\s+cash|extra\s+income|double\s+your\s+income|financial\s+freedom|eliminate\s+debt|get\s+out\s+of\s+debt|be\s+your\s+own\s+boss)\b/i,
    },
    {
        phrase: 'satisfaction guaranteed',
        category: 'Scam & Fraud',
        severity: 'high',
        suggestion: "Use 'demonstrated track record' or 'client references'",
        matcher: /\b(satisfaction\s+guaranteed|100%\s+guaranteed|money\s+back\s+guarantee|full\s+refund\s+guaranteed|risk[\s-]free\s+guarantee)\b/i,
    },
    {
        phrase: 'guaranteed returns',
        category: 'Scam & Fraud',
        severity: 'high',
        suggestion: "Avoid investment guarantee promises in outreach",
        matcher: /\b(guaranteed\s+(returns?|wealth|income|crypto|roi))\b/i,
    },
    {
        phrase: 'wire transfer',
        category: 'Scam & Fraud',
        severity: 'high',
        suggestion: "Financial fraud filters flag wire transfer and payment instructions",
        matcher: /\b(wire\s+transfer|western\s+union|moneygram|unclaimed\s+(funds|inheritance|lottery))\b/i,
    },
    {
        phrase: 'winner / prize',
        category: 'Scam & Fraud',
        severity: 'high',
        suggestion: "Remove lottery or contest prize terminology",
        matcher: /\b(you\s+(are\s+a\s+|have\s+been\s+selected\s+as\s+a\s+)?winner|winning\s+prize|claim\s+your\s+prize|instant\s+winner)\b/i,
    },
    {
        phrase: 'blacklisted pharma',
        category: 'Scam & Fraud',
        severity: 'high',
        suggestion: "Remove blacklisted pharmaceutical references",
        matcher: /\b(viagra|vicodin|cialis|valium|xanax)\b/i,
    },
    {
        phrase: 'gambling / casino',
        category: 'Scam & Fraud',
        severity: 'high',
        suggestion: "Remove gambling and casino terms",
        matcher: /\b(casino|jackpot|poker\s+bonus|online\s+betting|lottery\s+ticket)\b/i,
    },
    {
        phrase: 'passwords / credentials',
        category: 'Scam & Fraud',
        severity: 'high',
        suggestion: "Anti-phishing filters aggressively quarantine emails requesting credentials or passwords",
        matcher: /\b(passwords?|credentials?|social\s+security\s+number|pin\s+number)\b/i,
    },

    // --- 2. Urgency & High-Pressure Sales (Medium Risk) ---
    {
        phrase: 'act now',
        category: 'Urgency & Pressure',
        severity: 'medium',
        suggestion: "Use 'available this week' or 'timely opportunity'",
        matcher: /\b(act\s+now|order\s+now|buy\s+direct|call\s+now)\b/i,
    },
    {
        phrase: 'apply now',
        category: 'Urgency & Pressure',
        severity: 'medium',
        suggestion: "Use 'explore open positions' or 'view scope'",
        matcher: /\bapply\s+now\b/i,
    },
    {
        phrase: 'limited time offer',
        category: 'Urgency & Pressure',
        severity: 'medium',
        suggestion: "Use 'available this quarter' or 'scheduled roadmap'",
        matcher: /\b(limited\s+time(\s+offer)?|one\s+time\s+offer|special\s+promotion|once\s+in\s+a\s+lifetime|while\s+supplies\s+last)\b/i,
    },
    {
        phrase: 'open immediately',
        category: 'Urgency & Pressure',
        severity: 'high',
        suggestion: "Remove manufactured urgency; phishing filters flag 'open immediately'",
        matcher: /\b(open\s+immediately|read\s+immediately|urgent\s+response\s+required)\b/i,
    },
    {
        phrase: 'lowest price',
        category: 'Urgency & Pressure',
        severity: 'medium',
        suggestion: "Use 'competitive cost' or 'streamlined budget'",
        matcher: /\b(lowest\s+price|best\s+price|cheap\s+rates|discount\s+price)\b/i,
    },
    {
        phrase: 'no risk / no catch',
        category: 'Urgency & Pressure',
        severity: 'medium',
        suggestion: "Use 'low-friction conversation' or 'transparent trial'",
        matcher: /\b(no\s+risk|no\s+catch|no\s+gimmick|no\s+hidden\s+costs|no\s+obligation|no\s+purchase\s+necessary)\b/i,
    },

    // --- 3. Deceptive Phrasing & Deliverability Red Flags (Medium / High Risk) ---
    {
        phrase: 'not spam',
        category: 'Deceptive Phrasing',
        severity: 'high',
        suggestion: "Never state 'this is not spam' or 'not spam'; email gateways treat this phrase as a heuristic red flag",
        matcher: /\b(not\s+spam|this\s+is(n't|\s+not)\s+spam|unsolicited\s+email)\b/i,
    },
    {
        phrase: 'click here',
        category: 'Deliverability Flag',
        severity: 'medium',
        suggestion: "Use descriptive anchor text like 'review the case study' or 'view documentation' instead of 'click here'",
        matcher: /\b(click\s+here|click\s+below|click\s+now|click\s+this\s+link)\b/i,
    },
    {
        phrase: 'dear friend',
        category: 'Deliverability Flag',
        severity: 'medium',
        suggestion: "Use recipient's personalized name tag like 'Hi {fullName}' or 'Hi {firstName}'",
        matcher: /\bdear\s+friend\b/i,
    },
    {
        phrase: 'you have been selected',
        category: 'Deceptive Phrasing',
        severity: 'medium',
        suggestion: "Use 'reaching out based on your impressive background'",
        matcher: /\byou\s+have\s+been\s+selected\b/i,
        isBenign: (context: string) => {
            // Whitelist hiring context like "selected for an interview"
            return /\bselected\s+for\s+(an?\s+)?interview\b/i.test(context);
        }
    },
    {
        phrase: 'valuable',
        category: 'Deceptive Phrasing',
        severity: 'medium',
        suggestion: "Avoid commercial hype; professional phrasing like 'time is valuable' or 'valuable insights' is completely safe",
        matcher: /\bvaluable\b/i,
        isBenign: (context: string) => {
            // Safe when referring to time, insights, feedback, perspective
            return /\b(valuable\s+(time|insight|insights|feedback|input|experience|perspective|contribution)|time\s+is\s+valuable)\b/i.test(context);
        }
    }
];

/**
 * Strips HTML tags and normalizes whitespace
 */
function stripHtml(html: string): string {
    return html
        .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
        .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
        .replace(/<[^>]+>/g, ' ')
        .replace(/&nbsp;/gi, ' ')
        .replace(/&amp;/gi, '&')
        .replace(/&lt;/gi, '<')
        .replace(/&gt;/gi, '>')
        .replace(/\s+/g, ' ')
        .trim();
}

/**
 * Counts words in raw text
 */
function getWordCount(text: string): number {
    const cleaned = text.trim();
    if (!cleaned) return 0;
    return cleaned.split(/\s+/).filter(Boolean).length;
}

/**
 * Extracts links count and flags link shorteners
 */
function analyzeLinks(bodyHtml: string): { linkCount: number; hasSuspiciousShortener: boolean } {
    const anchorMatches = bodyHtml.match(/<a\s+[^>]*href=["']([^"']*)["'][^>]*>/gi) || [];
    const rawUrlMatches = bodyHtml.match(/https?:\/\/[^\s<"']+/gi) || [];
    const totalCount = Math.max(anchorMatches.length, rawUrlMatches.length);
    const hasSuspiciousShortener = SUSPICIOUS_SHORTENER_REGEX.test(bodyHtml);

    return {
        linkCount: totalCount,
        hasSuspiciousShortener,
    };
}

/**
 * Counts all-caps words in text while excluding allowable tech acronyms
 */
function countAllCapsWords(text: string): number {
    const words = text.split(/\s+/);
    let count = 0;
    for (const w of words) {
        const cleaned = w.replace(/[^A-Za-z]/g, '');
        if (cleaned.length >= 4 && cleaned === cleaned.toUpperCase() && !ALLOWED_ACRONYMS.has(cleaned)) {
            count++;
        }
    }
    return count;
}

/**
 * Analyzes an email subject and body for spam triggers and deliverability indicators
 */
export function analyzeSpamScore(subject: string, bodyHtml: string): SpamReport {
    let score = 100;
    const issues: SpamIssue[] = [];
    const detectedWords: DetectedSpamWord[] = [];
    const triggerWordsFound: string[] = [];

    const rawSubject = (subject || '').trim();
    const rawBody = stripHtml(bodyHtml || '');
    const combinedText = `${rawSubject} ${rawBody}`;

    const wordCount = getWordCount(rawBody);
    const readingTimeSeconds = Math.max(5, Math.round((wordCount / 200) * 60));
    const { linkCount, hasSuspiciousShortener } = analyzeLinks(bodyHtml || '');
    const allCapsCount = countAllCapsWords(rawBody);
    const hasPersonalization = /\{[a-zA-Z0-9_.-]+\}/i.test(`${rawSubject} ${rawBody}`);

    // --- 1. Subject Line Analysis ---
    if (!rawSubject) {
        score -= 30;
        issues.push({
            type: 'Subject Missing',
            message: 'Email has no subject line. Subject lines are critical for inbox placement.',
            severity: 'high',
            suggestion: 'Add a concise, descriptive subject line (e.g. Inquiry regarding {jobTitle} at {companyName}).',
        });
    } else {
        // Subject Length
        if (rawSubject.length < 8) {
            score -= 10;
            issues.push({
                type: 'Subject Too Short',
                message: 'Subject line is under 8 characters. Descriptive subjects get significantly higher open rates.',
                severity: 'medium',
                suggestion: 'Expand the subject line to 3-7 words clearly stating the email purpose.',
            });
        } else if (rawSubject.length > 65) {
            score -= 6;
            issues.push({
                type: 'Subject Too Long',
                message: 'Subject exceeds 65 characters and will be truncated on mobile mail clients.',
                severity: 'info',
                suggestion: 'Shorten subject line to under 50 characters for clean mobile previewing.',
            });
        }

        // All-Caps Subject Line
        if (rawSubject.length >= 8 && rawSubject === rawSubject.toUpperCase() && /[A-Z]/.test(rawSubject)) {
            score -= 25;
            issues.push({
                type: 'All-Caps Subject',
                message: 'Entire subject line is capitalized. This is one of the highest spam filter triggers.',
                severity: 'high',
                suggestion: 'Use sentence case or title case instead of all capital letters.',
            });
        }

        // Deceptive Reply/Forward Prefix
        if (/^(re:|fwd:)\s+/i.test(rawSubject)) {
            score -= 15;
            issues.push({
                type: 'Deceptive Prefix',
                message: 'Cold outreach starting with "Re:" or "Fwd:" triggers deceptive marketing heuristics.',
                severity: 'high',
                suggestion: 'Remove fake "Re:" prefixes from unsolicited cold emails.',
            });
        }

        // Excessive Punctuation in Subject
        if (/[!?]{2,}/.test(rawSubject)) {
            score -= 12;
            issues.push({
                type: 'Subject Punctuation',
                message: 'Multiple consecutive exclamation or question marks in subject line trigger spam filters.',
                severity: 'medium',
                suggestion: 'Limit punctuation to a single mark or remove entirely.',
            });
        }

        // Currency or Promo Symbols in Subject
        if (/[$€£%]/i.test(rawSubject)) {
            score -= 12;
            issues.push({
                type: 'Subject Promo Symbols',
                message: 'Currency ($/€/£) or discount symbols in subject lines trigger corporate promotional filters.',
                severity: 'medium',
                suggestion: 'Move financial and pricing details into the body text.',
            });
        }
    }

    // --- 2. Body Formatting & Links Analysis ---
    if (wordCount > 0 && wordCount < 25) {
        score -= 10;
        issues.push({
            type: 'Body Too Brief',
            message: 'Email body is under 25 words. Very short notes are frequently treated as suspicious ping messages.',
            severity: 'medium',
            suggestion: 'Add 1-2 concise sentences outlining who you are and your value proposition.',
        });
    } else if (wordCount > 220) {
        score -= 6;
        issues.push({
            type: 'Body Too Long',
            message: 'Cold email exceeds 220 words. Concise cold emails (50-125 words) generate up to 40% higher response rates.',
            severity: 'info',
            suggestion: 'Condense copy to focus on a single clear call-to-action.',
        });
    }

    if (linkCount > 4) {
        score -= 15;
        issues.push({
            type: 'Excessive Links',
            message: `Email contains ${linkCount} links. Cold emails with 4+ links are heavily diverted to Gmail Spam or Promotions tabs.`,
            severity: 'high',
            suggestion: 'Reduce links to 1 or 2 high-priority destinations.',
        });
    } else if (linkCount >= 3) {
        score -= 8;
        issues.push({
            type: 'High Link Count',
            message: `Email contains ${linkCount} links. Industry deliverability standard is 1-2 links maximum for cold outreach.`,
            severity: 'medium',
            suggestion: 'Keep only your primary portfolio or scheduling link.',
        });
    }

    if (hasSuspiciousShortener) {
        score -= 15;
        issues.push({
            type: 'URL Shortener Detected',
            message: 'Generic URL shorteners (e.g. bit.ly, tinyurl) are heavily flagged by security filters because they obscure the actual destination.',
            severity: 'high',
            suggestion: 'Use direct, transparent company or portfolio domain URLs.',
        });
    }

    if (/[!?]{3,}/.test(rawBody) || (rawBody.match(/[!]/g) || []).length > 4) {
        score -= 8;
        issues.push({
            type: 'Excessive Punctuation',
            message: 'Excessive exclamation marks in email body trigger heuristic spam classification.',
            severity: 'medium',
            suggestion: 'Keep tone professional by minimizing exclamation points.',
        });
    }

    if (/[$€£]{2,}/.test(combinedText)) {
        score -= 10;
        issues.push({
            type: 'Repeated Currency Symbols',
            message: 'Repeated currency symbols (e.g. "$$$") are flagged by corporate email filters.',
            severity: 'medium',
            suggestion: 'State figures in plain text (e.g. "$50,000" rather than "$$$").',
        });
    }

    if (allCapsCount >= 3) {
        score -= 8;
        issues.push({
            type: 'All-Caps Text',
            message: 'Multiple all-caps capitalized words detected in email body. Text shouting triggers deliverability filters.',
            severity: 'medium',
            suggestion: 'Use standard sentence casing for emphasis.',
        });
    }

    // --- 3. Context-Aware Spam Trigger Words Scan ---
    let spamWordPenalty = 0;

    for (const def of TRIGGER_DEFINITIONS) {
        const inSubject = def.matcher.test(rawSubject);
        const inBody = def.matcher.test(rawBody);

        if (!inSubject && !inBody) continue;

        // Check if benign in context
        const subjectBenign = inSubject && def.isBenign ? def.isBenign(rawSubject) : false;
        const bodyBenign = inBody && def.isBenign ? def.isBenign(rawBody) : false;

        const flaggedInSubject = inSubject && !subjectBenign;
        const flaggedInBody = inBody && !bodyBenign;

        if (flaggedInSubject || flaggedInBody) {
            const location: 'subject' | 'body' | 'both' =
                flaggedInSubject && flaggedInBody ? 'both' : flaggedInSubject ? 'subject' : 'body';

            const cleanPhrase = def.phrase.replace(/\s*\(.*?\)/, '');
            triggerWordsFound.push(cleanPhrase);

            detectedWords.push({
                word: cleanPhrase,
                category: def.category,
                severity: def.severity,
                suggestion: def.suggestion,
                location,
            });

            // Subject triggers carry 2x penalty
            const penalty = def.severity === 'high'
                ? (flaggedInSubject ? 16 : 10)
                : (flaggedInSubject ? 10 : 6);

            spamWordPenalty += penalty;
        }
    }

    // Cap total trigger word penalty to prevent negative runaway
    score -= Math.min(spamWordPenalty, 40);

    if (detectedWords.length > 0) {
        const highSeverityCount = detectedWords.filter(w => w.severity === 'high').length;
        const displayList = detectedWords.slice(0, 3).map(w => `"${w.word}"`).join(', ');
        const extra = detectedWords.length > 3 ? ` and ${detectedWords.length - 3} more` : '';

        issues.push({
            type: 'Flagged Keywords',
            message: `Detected flagged spam/promotional terms: ${displayList}${extra}.`,
            severity: highSeverityCount > 0 ? 'high' : 'medium',
            suggestion: detectedWords[0]?.suggestion,
        });
    }

    // --- 4. Personalization Check ---
    if (!hasPersonalization) {
        score -= 8;
        issues.push({
            type: 'No Personalization',
            message: 'No variable tags (like {fullName} or {companyName}) detected. Personalized cold outreach achieves 3x higher reply rates.',
            severity: 'info',
            suggestion: 'Insert dynamic variable tags from your recipient spreadsheet.',
        });
    } else {
        // Small deliverability confidence boost for personalized outreach
        score += 2;
    }

    // Clamp score strictly between 15 and 100
    const finalScore = Math.max(15, Math.min(100, score));

    let rating: 'Excellent' | 'Good' | 'Needs Improvement' = 'Excellent';
    if (finalScore < 70) {
        rating = 'Needs Improvement';
    } else if (finalScore < 85) {
        rating = 'Good';
    }

    const metrics: DeliverabilityMetrics = {
        wordCount,
        readingTimeSeconds,
        linkCount,
        subjectLength: rawSubject.length,
        hasPersonalization,
        allCapsCount,
    };

    return {
        score: finalScore,
        rating,
        triggerWords: triggerWordsFound,
        spamWordsFound: triggerWordsFound,
        detectedWords,
        metrics,
        issues,
    };
}
