import { Request } from 'express';

export interface AuthRequest extends Request {
    user?: {
        id: string;
        email: string;
        fullName: string;
    };
}

export interface RegisterDTO {
    email: string;
    password: string;
    fullName: string;
}

export interface LoginDTO {
    email: string;
    password: string;
}

export interface CreateCampaignDTO {
    name: string;
    subject: string;
    body: string;
    recipients: RecipientDTO[];
    scheduledTime?: Date;
    batchSize?: number;
    batchDelay?: number;
}

export interface RecipientDTO {
    fullName: string;
    email: string;
    companyName: string;
    jobTitle?: string;
}

export interface EmailCredentialDTO {
    email: string;
    appPassword: string;
    isDefault?: boolean;
}

export interface UpdateCampaignDTO {
    name?: string;
    status?: string;
    scheduledTime?: Date;
    batchSize?: number;
    batchDelay?: number;
}

export interface PaginationQuery {
    page?: number;
    limit?: number;
}

export interface CampaignQuery extends PaginationQuery {
    status?: string;
}

export interface RecipientQuery extends PaginationQuery {
    status?: string;
}
