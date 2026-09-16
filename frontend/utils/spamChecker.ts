import { SpamReport, SpamIssue } from '../types';

const SPAM_TRIGGER_WORDS: string[] = [
    '100% free', '100% satisfied', 'act now', 'apply now', 'bad credit', 'bargain',
    'be your own boss', 'best price', 'big bucks', 'billion dollars', 'bonus', 'buy direct',
    'call now', 'cancel at any time', 'cash bonus', 'casino', 'certified', 'cheap',
    'clearance', 'click here', 'click now', 'compare rates', 'congratulations', 'credit card offers',
    'dear friend', 'direct email', 'discount', 'double your income', 'earn extra cash', 'earn money',
    'eliminate debt', 'exclusive deal', 'extra income', 'fast cash', 'financial freedom', 'free consultation',
    'free gift', 'free info', 'free membership', 'free money', 'free sample', 'free trial', 'full refund',
    'get out of debt', 'get paid', 'giveaway', 'guaranteed', 'income from home', 'increase sales',
    'instant', 'investment', 'join millions', 'limited time', 'lowest price', 'make money', 'million dollars',
    'miracle', 'money back', 'no catch', 'no cost', 'no credit check', 'no experience', 'no fees',
    'no gimmick', 'no hidden costs', 'no obligation', 'no purchase necessary', 'no risk', 'not spam',
    'now only', 'one time offer', 'open immediately', 'order now', 'passwords', 'pennies a day',
    'potential earnings', 'prize', 'pure profit', 'refund', 'risk-free', 'save big', 'save money',
    'satisfaction guaranteed', 'special promotion', 'stop snoring', 'this isn\'t spam', 'unlimited',
    'unsolicited', 'urgent', 'valuable', 'viagra', 'vicodin', 'weight loss', 'while supplies last',
    'win', 'winner', 'winning', 'work from home', 'you have been selected'
];

/**
 * Strips HTML tags to extract raw visible text
 */
function stripHtml(html: string): string {
    return html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

/**
 * Analyzes an email subject and body for spam triggers and deliverability indicators
 */
export function analyzeSpamScore(subject: string, bodyHtml: string): SpamReport {
    let score = 100;
    const issues: SpamIssue[] = [];
    const triggerWordsFound: string[] = [];

    const rawSubject = (subject || '').trim();
    const rawBody = stripHtml(bodyHtml || '');
    const combinedText = `${rawSubject} ${rawBody}`.toLowerCase();

    // 1. Subject Length Check
    if (!rawSubject) {
        score -= 25;
        issues.push({
            type: 'Subject Missing',
            message: 'Email has no subject line. Subject lines are critical for inbox delivery.',
            severity: 'high',
        });
    } else if (rawSubject.length < 8) {
        score -= 10;
        issues.push({
            type: 'Subject Too Short',
            message: 'Subject line is very short (under 8 characters). Clear, descriptive subjects get higher open rates.',
            severity: 'medium',
        });
    } else if (rawSubject.length > 70) {
        score -= 8;
        issues.push({
            type: 'Subject Too Long',
            message: 'Subject exceeds 70 characters and will likely be truncated on mobile email clients.',
            severity: 'info',
        });
    }

    // 2. All-Caps Subject Line
    if (rawSubject.length >= 8 && rawSubject === rawSubject.toUpperCase() && /[A-Z]/.test(rawSubject)) {
        score -= 20;
        issues.push({
            type: 'All-Caps Subject',
            message: 'Entire subject line is capitalized. This is one of the highest spam filter flags.',
            severity: 'high',
        });
    }

    // 3. Repeated Punctuation Check
    if (/[!?]{2,}/.test(rawSubject) || /[!?]{3,}/.test(rawBody)) {
        score -= 10;
        issues.push({
            type: 'Excessive Punctuation',
            message: 'Multiple consecutive exclamation or question marks (e.g. "!!" or "???") trigger heuristic filters.',
            severity: 'medium',
        });
    }

    // 4. Repeated Currency Symbols
    if (/[$€£]{2,}/.test(combinedText)) {
        score -= 10;
        issues.push({
            type: 'Excessive Currency Symbols',
            message: 'Repeated currency symbols (e.g. "$$$") are flagged by corporate spam gateways.',
            severity: 'medium',
        });
    }

    // 5. Spam Trigger Words Scan
    for (const phrase of SPAM_TRIGGER_WORDS) {
        const regex = new RegExp(`\\b${phrase.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&')}\\b`, 'i');
        if (regex.test(combinedText)) {
            triggerWordsFound.push(phrase);
            score -= 8;
        }
    }

    if (triggerWordsFound.length > 0) {
        const displayList = triggerWordsFound.slice(0, 4).map(w => `"${w}"`).join(', ');
        const extra = triggerWordsFound.length > 4 ? ` and ${triggerWordsFound.length - 4} more` : '';
        issues.push({
            type: 'Spam Trigger Words',
            message: `Detected flagged spam keywords: ${displayList}${extra}. Consider substituting with natural professional phrasing.`,
            severity: triggerWordsFound.length >= 3 ? 'high' : 'medium',
        });
    }

    // 6. Personalization Tags Check
    const hasPersonalization = /{(?:fullName|name|firstName|lastName|companyName|company|jobTitle|title)}/i.test(`${rawSubject} ${rawBody}`);
    if (!hasPersonalization) {
        score -= 8;
        issues.push({
            type: 'No Personalization Tags',
            message: 'No variable tags (like {fullName} or {companyName}) detected. Personalized outreach receives up to 3x higher reply rates.',
            severity: 'info',
        });
    }

    // 7. Body Length Check
    if (rawBody.length > 0 && rawBody.length < 50) {
        score -= 10;
        issues.push({
            type: 'Body Too Brief',
            message: 'Email body is under 50 characters. Very short messages are often treated as test emails or suspicious pings.',
            severity: 'medium',
        });
    }

    // Clamp score
    const finalScore = Math.max(15, Math.min(100, score));

    let rating: 'Excellent' | 'Good' | 'Needs Improvement' = 'Excellent';
    if (finalScore < 75) {
        rating = 'Needs Improvement';
    } else if (finalScore < 90) {
        rating = 'Good';
    }

    return {
        score: finalScore,
        rating,
        triggerWords: triggerWordsFound,
        spamWordsFound: triggerWordsFound,
        issues,
    };
}
