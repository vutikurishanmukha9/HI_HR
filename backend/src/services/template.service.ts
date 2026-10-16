import prisma from '../config/database';
import { AppError } from '../middleware/errorHandler';

export interface TemplateDTO {
    name: string;
    category?: string;
    subject: string;
    body: string;
}

export interface TemplateRecord {
    id: string;
    name: string;
    category: string;
    subject: string;
    body: string;
    isSystem: boolean;
    createdAt?: Date;
    updatedAt?: Date;
}

// Built-in battle-tested starter templates across 7 specialized outreach categories
const SYSTEM_TEMPLATES: TemplateRecord[] = [
    // --- 1. Cold Outreach (Job Seekers & Technical Talent) ---
    {
        id: 'sys-exec-inquiry',
        name: 'Direct Opportunity Inquiry',
        category: 'Cold Outreach',
        subject: 'Inquiry: {jobTitle} opportunities at {companyName}',
        body: `<p>Hi {fullName},</p><p>I have been closely following {companyName}'s work and recent milestones. As an experienced professional with a background in delivering high-impact initiatives, I wanted to reach out directly to inquire about opportunities within your team.</p><p>I would welcome the opportunity to discuss how my skill set aligns with your upcoming goals. Do you have 10 minutes next week for a brief introductory conversation?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-value-pitch',
        name: 'The Value-First Pitch (Problem Solver)',
        category: 'Cold Outreach',
        subject: "Quick idea regarding {companyName}'s growth & {jobTitle} role",
        body: `<p>Hi {fullName},</p><p>I have been studying {companyName}'s recent product direction and noticed how rapidly your team is scaling its core initiatives. Based on my experience driving similar projects, I developed a few actionable observations on how to optimize key workflows for this stage.</p><p>I would be glad to share these thoughts with you or your team, no strings attached. Would you be open to a 10-minute chat this Thursday or Friday?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-recent-milestone',
        name: 'Recent Milestone & Team Scaling',
        category: 'Cold Outreach',
        subject: "Congrats on {companyName}'s milestone + quick question",
        body: `<p>Hi {fullName},</p><p>Congratulations on {companyName}'s recent announcements and growth milestones. Seeing your progress in the industry has been truly impressive.</p><p>With this expansion, I imagine building out your {jobTitle} capabilities is a top priority. My background centers on scaling technical systems and accelerating team deliverables under fast-paced conditions.</p><p>Are you currently open to conversations with experienced talent looking to contribute to your next phase?</p><p>Kind regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-portfolio-showcase',
        name: 'Technical Competency & Case Study',
        category: 'Cold Outreach',
        subject: '{fullName} - {jobTitle} portfolio & case study for {companyName}',
        body: `<p>Hi {fullName},</p><p>Rather than sending a generic resume, I wanted to share a concrete example of work that closely parallels what {companyName} is building.</p><p>In my recent role, I led an initiative that improved system performance by over 35% while streamlining collaborative developer workflows. I believe this experience translates directly to the challenges your {jobTitle} team navigates daily.</p><p>I have attached a brief one-page overview for your review. Would you be available for a brief introductory call next week?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-alumni-network',
        name: 'Alumni & Shared Network Connection',
        category: 'Cold Outreach',
        subject: 'Fellow alum reaching out regarding {companyName} & {jobTitle}',
        body: `<p>Hi {fullName},</p><p>I noticed your impressive trajectory at {companyName} and wanted to reach out as a fellow member of our shared professional and academic community.</p><p>I am exploring new challenges in the {jobTitle} space and greatly admire the standard of excellence your team maintains. If your schedule permits, I would appreciate the chance to hear your perspective on the company's trajectory and team culture.</p><p>Would you have 10 minutes for a brief introductory chat sometime next week?</p><p>Warm regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-opensource-pitch',
        name: 'Open-Source & Code Contribution Pitch',
        category: 'Cold Outreach',
        subject: 'Contributions to open source & ideas for {companyName}',
        body: `<p>Hi {fullName},</p><p>I have been following your engineering blogs and open-source contributions at {companyName}. I recently worked on a public module addressing similar architectural patterns and thought you might find the approach interesting.</p><p>As an active builder in this space, I am eager to apply this hands-on depth to your core initiatives as a {jobTitle}.</p><p>Would you be open to a 10-minute technical exchange sometime this week?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-competitor-experience',
        name: 'Direct Industry Domain Expertise',
        category: 'Cold Outreach',
        subject: '{jobTitle} insights from direct industry experience for {companyName}',
        body: `<p>Hi {fullName},</p><p>Having worked directly in our industry over the past few years, I have experienced first-hand the operational hurdles and competitive dynamics that {companyName} is navigating.</p><p>Because of this shared context, I can step into your {jobTitle} opening with zero ramp-up time and deliver meaningful impact from day one.</p><p>Could we find 10 minutes for an exploratory chat this Thursday or Friday?</p><p>Kind regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-fast-learner-entry',
        name: 'High-Velocity Builder (Hunger & Speed)',
        category: 'Cold Outreach',
        subject: 'Driven {jobTitle} applicant for {companyName}',
        body: `<p>Hi {fullName},</p><p>I know your inbox is full, so I will be direct: I am a disciplined, high-velocity builder looking to earn a spot on {companyName}'s engineering team.</p><p>Over the past year, I have shipped multiple complete projects end-to-end, prioritizing clean code, test coverage, and documentation. I am ready to tackle whatever high-priority backlog items your team needs executed.</p><p>Would you be open to giving me 10 minutes to prove what I can bring to your team?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-product-ux-audit',
        name: 'Product & UX Usability Observation',
        category: 'Cold Outreach',
        subject: '3 quick observations on {companyName}\'s user experience',
        body: `<p>Hi {fullName},</p><p>I spent time testing {companyName}'s product this week and took notes on 3 high-impact friction points where minor adjustments could meaningfully boost user activation.</p><p>I put together a short summary with wireframe suggestions that I would be glad to share with your product and engineering leadership.</p><p>Would you be interested in taking a look, completely free with no obligations?</p><p>Kind regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-contract-to-hire',
        name: 'Low-Risk 2-Week Trial Project',
        category: 'Cold Outreach',
        subject: 'Low-risk 2-week trial project for {companyName} ({jobTitle})',
        body: `<p>Hi {fullName},</p><p>I understand hiring decisions carry risk and take valuable bandwidth away from core development. To eliminate that risk, I would like to propose a 2-week scoped trial project for {companyName}.</p><p>Let me tackle a specific feature or backlog issue on your roadmap. If you are thrilled with my work, we can discuss a permanent {jobTitle} engagement; if not, you keep the code and we part ways amicably.</p><p>Would you be open to discussing a candidate project this week?</p><p>Best regards,</p>`,
        isSystem: true,
    },

    // --- 2. Sales & Business Development (B2B Outreach) ---
    {
        id: 'sys-b2b-funding-trigger',
        name: 'Post-Funding / Scaling Milestone Outreach',
        category: 'Business Development',
        subject: 'Scaling {companyName}\'s tech operations post-announcement',
        body: `<p>Hi {fullName},</p><p>Congratulations on {companyName}'s recent milestones and team growth. As companies scale rapidly through this stage, internal operational bottlenecks and infrastructure friction typically become key challenges.</p><p>We help fast-growing organizations streamline their outbound and operational pipelines, increasing execution velocity by over 35% without adding headcount.</p><p>Do you have 10 minutes next Tuesday to see if this could accelerate your current roadmap?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-b2b-case-study',
        name: 'Proven Peer Case Study / ROI Benchmark',
        category: 'Business Development',
        subject: 'How similar teams increased pipeline output by 40%',
        body: `<p>Hi {fullName},</p><p>We recently partnered with a high-growth team facing similar market challenges to {companyName}. Within 60 days, they reduced pipeline delays by 40% while cutting tooling overhead.</p><p>Given your focus on delivering results as {jobTitle}, I thought these benchmark figures would be directly relevant to your quarterly targets.</p><p>Would you be open to reviewing a brief 2-minute summary of how they achieved this?</p><p>Kind regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-b2b-short-ceo',
        name: 'The 3-Sentence Executive Pitch',
        category: 'Business Development',
        subject: 'Quick question regarding {companyName}\'s workflow',
        body: `<p>Hi {fullName},</p><p>We built a streamlined solution that helps organizations like {companyName} automate repetitive communication and guarantee 99%+ deliverability.</p><p>Teams using our engine have saved an average of 15 hours per week while improving conversion rates.</p><p>Are you open to a brief 5-minute chat this Thursday or Friday?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-b2b-video-audit',
        name: 'Complimentary 90-Second Custom Teardown',
        category: 'Business Development',
        subject: 'Recorded a 90-second teardown for {companyName}',
        body: `<p>Hi {fullName},</p><p>Rather than sending a generic cold pitch, I recorded a quick 90-second video walkthrough showing 2 specific opportunities to optimize your current {jobTitle} workflow at {companyName}.</p><p>It highlights concrete improvements you can implement immediately, with or without us.</p><p>Would you like me to send the private link over for your review?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-b2b-cost-efficiency',
        name: 'Operational Efficiency & Cost Reduction',
        category: 'Business Development',
        subject: 'Reducing operational overhead for {companyName}',
        body: `<p>Hi {fullName},</p><p>In today's market, maximizing team efficiency while eliminating redundant SaaS subscriptions is top of mind for forward-thinking leadership.</p><p>Our infrastructure consolidates disparate outbound and validation tooling into a unified, high-performance platform, typically reducing tool spend by 30-50%.</p><p>Would you be interested in seeing a quick cost comparison for {companyName}?</p><p>Kind regards,</p>`,
        isSystem: true,
    },

    // --- 3. Follow-Up Sequences ---
    {
        id: 'sys-polite-followup-1',
        name: 'Polite Follow-up #1 (Gentle Bump)',
        category: 'Follow Up',
        subject: 'Re: Inquiry regarding opportunities at {companyName}',
        body: `<p>Hi {fullName},</p><p>I wanted to follow up briefly on my previous note regarding opportunities at {companyName}. I recognize you have a demanding schedule, so I wanted to bring this back to your attention in case you had a moment to review.</p><p>I remain very interested in contributing to your team's goals. Would you have a few minutes for a brief conversation this week?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-followup-value-add',
        name: 'Follow-up #2: Adding New Insight',
        category: 'Follow Up',
        subject: 'Quick update regarding {companyName} & {jobTitle}',
        body: `<p>Hi {fullName},</p><p>Following up on my previous message, I recently came across an industry analysis directly relevant to the market {companyName} is leading, and it immediately reminded me of your team's approach.</p><p>Given my experience delivering measurable outcomes in similar environments, I would welcome the chance to share how I could hit the ground running as your next {jobTitle}.</p><p>Would 10 minutes next Tuesday or Wednesday work for a quick conversation?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-followup-quick-thought',
        name: 'Follow-up #3: "Did this get buried?" (Short)',
        category: 'Follow Up',
        subject: 'Re: {companyName} {jobTitle} inquiry',
        body: `<p>Hi {fullName},</p><p>I know how quickly inboxes fill up. Just wanted to float this to the top of your list in case it got buried under other priorities.</p><p>Would you have 5 minutes for a quick chat regarding your {jobTitle} needs at {companyName}?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-followup-resource-share',
        name: 'Follow-up #4: Helpful Industry Resource',
        category: 'Follow Up',
        subject: 'Thought of {companyName}: relevant industry benchmark',
        body: `<p>Hi {fullName},</p><p>Regardless of whether you are actively hiring for {jobTitle} roles right now, I came across this benchmark report on technical delivery efficiency and thought your team at {companyName} would find value in it.</p><p>No need to reply if your plate is full. Hope you find the insights helpful!</p><p>Warm regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-breakup-email',
        name: 'Follow-up #5: Permission to Close the Loop',
        category: 'Follow Up',
        subject: 'Closing the loop on {companyName} {jobTitle}',
        body: `<p>Hi {fullName},</p><p>I have reached out a couple of times regarding {jobTitle} opportunities with {companyName}. Since I have not heard back, I assume priorities have shifted or your hiring roadmap is currently full.</p><p>I will stop following up for now so I don't clutter your inbox. If timing aligns better down the road, please feel free to reconnect.</p><p>Wishing you and {companyName} continued success,</p>`,
        isSystem: true,
    },

    // --- 4. Referral & Professional Networking ---
    {
        id: 'sys-referral-request',
        name: 'Team Introduction & Referral Request',
        category: 'Referral',
        subject: 'Connecting regarding {companyName} & team roles',
        body: `<p>Hi {fullName},</p><p>I noticed your work at {companyName} and wanted to reach out. I am currently exploring new career opportunities and admire the high-caliber culture your team has established.</p><p>I noticed an open {jobTitle} position and would love to ask you a couple of brief questions about your experience on the team before submitting an application.</p><p>Would you be open to a quick 5-minute conversation or message exchange?</p><p>Thank you for your time,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-coffee-chat',
        name: 'Informational Coffee Chat (Low Friction)',
        category: 'Referral',
        subject: '10-min virtual coffee chat: {companyName} {jobTitle} journey',
        body: `<p>Hi {fullName},</p><p>I have been following your career journey and contributions at {companyName}. As someone currently working towards similar career milestones in the {jobTitle} domain, your experience stands out as a benchmark.</p><p>I would be grateful for the chance to ask 2 or 3 quick questions about your experience transitioning into your current scope. I promise to be concise and respect your time.</p><p>Would you be open to a quick 10-minute virtual coffee over Zoom or Google Meet?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-mutual-connection',
        name: 'Mutual Connection Introduction Bridge',
        category: 'Referral',
        subject: 'Connecting through shared network regarding {companyName}',
        body: `<p>Hi {fullName},</p><p>I noticed we share mutual connections in the industry, and your contributions to {companyName} frequently came up as a benchmark for high-execution teams.</p><p>I am exploring senior {jobTitle} roles where I can contribute to strong engineering cultures. If you have a moment, I would value any brief perspective on how your team approaches technical scaling.</p><p>Would you be open to connecting?</p><p>Warm regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-event-speaker',
        name: 'Conference & Event Speaker Follow-up',
        category: 'Referral',
        subject: 'Loved your talk on architecture + quick question',
        body: `<p>Hi {fullName},</p><p>I recently attended your presentation regarding technical scaling and was particularly struck by your insights on system modularity and team ownership.</p><p>Seeing how you implement these principles at {companyName} made it clear that your organization is building for the long term. As an experienced {jobTitle}, I would welcome the chance to connect and stay in touch.</p><p>Thank you for sharing your expertise with the community!</p><p>Best regards,</p>`,
        isSystem: true,
    },

    // --- 5. Executive & Leadership Outreach ---
    {
        id: 'sys-executive-alignment',
        name: 'Direct to Hiring Lead / Department Head',
        category: 'Executive',
        subject: "Supporting your team's upcoming roadmap at {companyName}",
        body: `<p>Hi {fullName},</p><p>I am reaching out to you directly because I know leadership at {companyName} prioritizes high-execution standards and clear technical accountability.</p><p>With a proven track record of spearheading complex projects from inception to production, I am confident in my ability to step in as your {jobTitle} and drive immediate measurable momentum.</p><p>Could we set up a brief 10-minute intro call to discuss your strategic priorities for this upcoming quarter?</p><p>Respectfully,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-two-initiatives',
        name: 'The "Two Initiatives" Leadership Pitch',
        category: 'Executive',
        subject: "Thoughts on {companyName}'s {jobTitle} priorities for this quarter",
        body: `<p>Hi {fullName},</p><p>Knowing that your time is valuable, I will get straight to the point. If brought on board as your next {jobTitle} at {companyName}, here are two specific areas where I can deliver immediate impact:</p><ol><li><strong>Process Velocity:</strong> Streamlining delivery pipelines to accelerate feature releases without compromising quality.</li><li><strong>Scalable Architecture:</strong> Reinforcing core reliability standards to ensure smooth operations through your next phase of growth.</li></ol><p>I would welcome 10 minutes to discuss whether these priorities align with your current roadmap.</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-exec-turnaround',
        name: 'Technical Debt & Velocity Modernization',
        category: 'Executive',
        subject: 'Accelerating technical velocity at {companyName}',
        body: `<p>Hi {fullName},</p><p>As engineering organizations grow, maintaining deployment velocity while managing legacy technical debt becomes one of leadership\'s most difficult balancing acts.</p><p>In my recent roles, I have specialized in identifying deployment bottlenecks, modernizing CI/CD infrastructure, and cutting release cycles by half without disrupting production uptime.</p><p>Would 10 minutes next week be worthwhile to explore how I could support {companyName}\'s upcoming release goals?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-exec-fractional',
        name: 'Fractional / Advisory Leadership Offer',
        category: 'Executive',
        subject: 'Fractional {jobTitle} advisory for {companyName}',
        body: `<p>Hi {fullName},</p><p>Many fast-scaling companies need senior technical direction but do not necessarily require a full-time executive hire at their current stage.</p><p>I provide fractional {jobTitle} advisory (10-15 hours/week), focusing on technical architecture, hiring evaluations, and engineering roadmap alignment at a fraction of full-time cost.</p><p>Would you be open to an exploratory conversation to see if this model fits your upcoming objectives?</p><p>Kind regards,</p>`,
        isSystem: true,
    },

    // --- 6. Recruiter & Talent Sourcing ---
    {
        id: 'sys-passive-candidate',
        name: 'Passive Talent Sourcing (High-Demand Role)',
        category: 'Recruiter',
        subject: 'Exciting {jobTitle} opportunity tailored to your background',
        body: `<p>Hi {fullName},</p><p>I came across your profile and was thoroughly impressed by your background and achievements in the industry. Your track record of delivering high-quality solutions caught our attention.</p><p>We are currently building out our leadership in the {jobTitle} space at {companyName}, and your experience aligns remarkably well with the technical challenges and autonomy this role offers.</p><p>Even if you are not actively looking, would you be open to an informal 10-minute exploratory conversation this week?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-fasttrack-interview',
        name: 'Fast-Track Interview Invitation',
        category: 'Recruiter',
        subject: 'Fast-track interview invitation for {jobTitle}',
        body: `<p>Hi {fullName},</p><p>Our engineering and leadership team at {companyName} recently reviewed your portfolio and public contributions, and we believe you would be an exceptional addition to our organization.</p><p>We are looking to fill a pivotal {jobTitle} position and would like to fast-track your candidacy, skipping introductory screening steps to connect you directly with our hiring team.</p><p>Are you open to reviewing the role overview and scheduling a brief introductory discussion?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-executive-search',
        name: 'Confidential Leadership Search',
        category: 'Recruiter',
        subject: 'Confidential Leadership Search: {jobTitle} at {companyName}',
        body: `<p>Hi {fullName},</p><p>I am conducting a targeted search for an executive {jobTitle} role on behalf of {companyName}. Given your strategic accomplishments and reputation in the sector, your profile stood out immediately.</p><p>This role carries significant organizational influence, competitive executive equity, and direct strategic oversight over critical company initiatives.</p><p>Would you be open to a confidential conversation to explore whether this opportunity aligns with your future ambitions?</p><p>Kind regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-remote-flexibility',
        name: 'Remote-First & Async Autonomy Sourcing',
        category: 'Recruiter',
        subject: 'Remote {jobTitle} opportunity with high autonomy at {companyName}',
        body: `<p>Hi {fullName},</p><p>I noticed your background in building resilient software systems and wanted to reach out regarding a high-impact {jobTitle} opening with our team at {companyName}.</p><p>We operate as a 100% remote, async-first organization with zero micromanagement, top-tier compensation, and deep technical challenges.</p><p>If you are open to learning more, could I send over a 2-page overview of the role and team compensation philosophy?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-re-engage-talent',
        name: 'Re-Engaging Previous High-Caliber Talent',
        category: 'Recruiter',
        subject: 'Checking back in: new {jobTitle} scope at {companyName}',
        body: `<p>Hi {fullName},</p><p>It has been a little while since we last spoke, but our team at {companyName} has continued to follow your work with great admiration.</p><p>We recently opened an expanded {jobTitle} role that maps closely to the exact strategic direction you expressed interest in during our previous discussions.</p><p>Would you be open to catching up for 10 minutes to see what has evolved on our end?</p><p>Warm regards,</p>`,
        isSystem: true,
    },

    // --- 7. Consulting & Agency Services ---
    {
        id: 'sys-agency-audit',
        name: 'Complimentary Architecture / Security Audit',
        category: 'Consulting',
        subject: 'Complimentary audit for {companyName}\'s web performance',
        body: `<p>Hi {fullName},</p><p>Our technical team conducted a brief external performance and deliverability audit of {companyName}'s web assets and identified 3 key areas where latency and domain reputation could be improved.</p><p>We compiled our findings into a concise 1-page report detailing actionable fixes you can implement right away.</p><p>Would you like me to send the PDF over for your engineering team to review?</p><p>Best regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-agency-bandwidth',
        name: 'On-Demand Engineering Surge Bandwidth',
        category: 'Consulting',
        subject: 'Additional {jobTitle} engineering bandwidth for {companyName}',
        body: `<p>Hi {fullName},</p><p>Hitting tight product deadlines often requires temporary technical reinforcement without the overhead and long delays of full-time hiring.</p><p>We provide senior, vetted {jobTitle} engineers who embed directly into your sprints within 48 hours to help accelerate delivery and clear technical roadblocks.</p><p>Would you be open to a brief chat to see how we could support your upcoming release sprint?</p><p>Kind regards,</p>`,
        isSystem: true,
    },
    {
        id: 'sys-agency-conversion',
        name: 'Conversion Rate & Funnel Optimization',
        category: 'Consulting',
        subject: 'Quick observations on {companyName}\'s user conversion flow',
        body: `<p>Hi {fullName},</p><p>While reviewing {companyName}'s current customer onboarding journey, I noticed a couple of minor usability friction points that may be causing avoidable user drop-off.</p><p>In similar client engagements, fixing these specific flow patterns resulted in an immediate 15-25% lift in onboarding conversions.</p><p>Would you be open to a quick 10-minute exchange to discuss these findings?</p><p>Best regards,</p>`,
        isSystem: true,
    }
];

export class TemplateService {
    async getTemplates(userId: string): Promise<TemplateRecord[]> {
        const userTemplates = await prisma.emailTemplate.findMany({
            where: { userId },
            orderBy: { updatedAt: 'desc' },
        });

        const formattedUserTemplates: TemplateRecord[] = userTemplates.map(t => ({
            id: t.id,
            name: t.name,
            category: t.category,
            subject: t.subject,
            body: t.body,
            isSystem: false,
            createdAt: t.createdAt,
            updatedAt: t.updatedAt,
        }));

        // Return user templates first, followed by system templates
        return [...formattedUserTemplates, ...SYSTEM_TEMPLATES];
    }

    async createTemplate(userId: string, data: TemplateDTO): Promise<TemplateRecord> {
        if (!data.name || !data.subject || !data.body) {
            throw new AppError('Template name, subject, and body are required', 400);
        }

        const created = await prisma.emailTemplate.create({
            data: {
                userId,
                name: data.name.trim(),
                category: (data.category || 'Cold Outreach').trim(),
                subject: data.subject.trim(),
                body: data.body,
            },
        });

        return {
            id: created.id,
            name: created.name,
            category: created.category,
            subject: created.subject,
            body: created.body,
            isSystem: false,
            createdAt: created.createdAt,
            updatedAt: created.updatedAt,
        };
    }

    async updateTemplate(userId: string, id: string, data: Partial<TemplateDTO>): Promise<TemplateRecord> {
        if (id.startsWith('sys-')) {
            throw new AppError('System templates cannot be edited. Please save as a new custom template.', 400);
        }

        const existing = await prisma.emailTemplate.findFirst({
            where: { id, userId },
        });

        if (!existing) {
            throw new AppError('Template not found', 404);
        }

        const updated = await prisma.emailTemplate.update({
            where: { id },
            data: {
                ...(data.name && { name: data.name.trim() }),
                ...(data.category && { category: data.category.trim() }),
                ...(data.subject && { subject: data.subject.trim() }),
                ...(data.body && { body: data.body }),
            },
        });

        return {
            id: updated.id,
            name: updated.name,
            category: updated.category,
            subject: updated.subject,
            body: updated.body,
            isSystem: false,
            createdAt: updated.createdAt,
            updatedAt: updated.updatedAt,
        };
    }

    async deleteTemplate(userId: string, id: string): Promise<{ success: boolean; message: string }> {
        if (id.startsWith('sys-')) {
            throw new AppError('System templates cannot be deleted', 400);
        }

        const existing = await prisma.emailTemplate.findFirst({
            where: { id, userId },
        });

        if (!existing) {
            throw new AppError('Template not found', 404);
        }

        await prisma.emailTemplate.delete({
            where: { id },
        });

        return { success: true, message: 'Template deleted successfully' };
    }
}

export default new TemplateService();
