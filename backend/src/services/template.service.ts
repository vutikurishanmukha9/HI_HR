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

// Built-in battle-tested starter templates
const SYSTEM_TEMPLATES: TemplateRecord[] = [
    // --- Cold Outreach (Job Seekers & Business Development) ---
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

    // --- Follow-Up Sequences ---
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
        id: 'sys-breakup-email',
        name: 'Follow-up #3: Permission to Close the Loop',
        category: 'Follow Up',
        subject: 'Closing the loop on {companyName} {jobTitle}',
        body: `<p>Hi {fullName},</p><p>I have reached out a couple of times regarding {jobTitle} opportunities with {companyName}. Since I have not heard back, I assume priorities have shifted or your hiring roadmap is currently full.</p><p>I will stop following up for now so I don't clutter your inbox. If timing aligns better down the road, please feel free to reconnect.</p><p>Wishing you and {companyName} continued success,</p>`,
        isSystem: true,
    },

    // --- Referral & Networking ---
    {
        id: 'sys-referral-request',
        name: 'Team Introduction & Referral',
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

    // --- Executive & Leadership ---
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

    // --- Recruiter & Talent Sourcing ---
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
