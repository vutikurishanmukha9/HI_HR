import prisma from '../config/database';
import { AppError } from '../middleware/errorHandler';
import { CreateCampaignDTO, UpdateCampaignDTO, CampaignQuery, RecipientDTO } from '../types';
import { removeDuplicateRecipients } from '../utils/excel';

export class CampaignService {
    async createCampaign(userId: string, data: CreateCampaignDTO, recipients: RecipientDTO[]) {
        // Remove duplicate recipients
        const uniqueRecipients = removeDuplicateRecipients(recipients);

        // Create campaign with recipients
        const campaign = await prisma.campaign.create({
            data: {
                userId,
                name: data.name,
                subject: data.subject,
                body: data.body,
                scheduledTime: data.scheduledTime,
                batchSize: data.batchSize || 10,
                batchDelay: data.batchDelay || 60,
                totalRecipients: uniqueRecipients.length,
                recipients: {
                    create: uniqueRecipients.map((recipient) => ({
                        fullName: recipient.fullName,
                        email: recipient.email,
                        companyName: recipient.companyName,
                        jobTitle: recipient.jobTitle,
                    })),
                },
            },
            include: {
                recipients: true,
            },
        });

        return campaign;
    }

    async getCampaigns(userId: string, query: CampaignQuery) {
        const page = query.page || 1;
        const limit = query.limit || 10;
        const skip = (page - 1) * limit;

        const where: any = { userId };
        if (query.status) {
            where.status = query.status;
        }

        const [campaigns, total] = await Promise.all([
            prisma.campaign.findMany({
                where,
                skip,
                take: limit,
                orderBy: { createdAt: 'desc' },
                select: {
                    id: true,
                    name: true,
                    subject: true,
                    status: true,
                    scheduledTime: true,
                    totalRecipients: true,
                    sentCount: true,
                    failedCount: true,
                    createdAt: true,
                    updatedAt: true,
                },
            }),
            prisma.campaign.count({ where }),
        ]);

        return {
            campaigns,
            total,
            page,
            totalPages: Math.ceil(total / limit),
        };
    }

    async getCampaignById(userId: string, campaignId: string) {
        const campaign = await prisma.campaign.findFirst({
            where: { id: campaignId, userId },
            include: {
                recipients: {
                    orderBy: { createdAt: 'asc' },
                },
                attachments: true,
            },
        });

        if (!campaign) {
            throw new AppError('Campaign not found', 404);
        }

        return campaign;
    }

    async updateCampaign(userId: string, campaignId: string, data: UpdateCampaignDTO) {
        // Check if campaign exists and belongs to user
        const existing = await prisma.campaign.findFirst({
            where: { id: campaignId, userId },
        });

        if (!existing) {
            throw new AppError('Campaign not found', 404);
        }

        // Don't allow updating if campaign is sending
        if (existing.status === 'sending') {
            throw new AppError('Cannot update campaign while sending', 400);
        }

        const campaign = await prisma.campaign.update({
            where: { id: campaignId },
            data: {
                ...data,
                updatedAt: new Date(),
            },
        });

        return campaign;
    }

    async deleteCampaign(userId: string, campaignId: string) {
        const campaign = await prisma.campaign.findFirst({
            where: { id: campaignId, userId },
        });

        if (!campaign) {
            throw new AppError('Campaign not found', 404);
        }

        // Don't allow deleting if campaign is sending
        if (campaign.status === 'sending') {
            throw new AppError('Cannot delete campaign while sending', 400);
        }

        await prisma.campaign.delete({
            where: { id: campaignId },
        });

        return { message: 'Campaign deleted successfully' };
    }

    async updateRecipientStatus(
        recipientId: string,
        status: string,
        errorMessage?: string
    ) {
        await prisma.recipient.update({
            where: { id: recipientId },
            data: {
                status,
                errorMessage,
                sentAt: status === 'sent' ? new Date() : undefined,
            },
        });
    }

    async updateCampaignCounts(campaignId: string) {
        const counts = await prisma.recipient.groupBy({
            by: ['status'],
            where: { campaignId },
            _count: true,
        });

        const sentCount = counts.find((c) => c.status === 'sent')?._count || 0;
        const failedCount = counts.find((c) => c.status === 'failed')?._count || 0;

        await prisma.campaign.update({
            where: { id: campaignId },
            data: { sentCount, failedCount },
        });
    }
}

export default new CampaignService();
