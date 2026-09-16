import nodemailer from 'nodemailer';
import prisma from '../config/database';
import { encrypt, decrypt } from '../utils/encryption';
import { AppError } from '../middleware/errorHandler';
import { EmailCredentialDTO } from '../types';

export class CredentialService {
    async createCredential(userId: string, data: EmailCredentialDTO) {
        const cleanEmail = data.email.trim().toLowerCase();
        const cleanPassword = (data.appPassword || '').replace(/\s+/g, '');

        // Verify with Gmail SMTP before saving to ensure credentials are valid
        const transporter = nodemailer.createTransport({
            host: 'smtp.gmail.com',
            port: 587,
            secure: false,
            auth: {
                user: cleanEmail,
                pass: cleanPassword,
            },
        });

        try {
            await transporter.verify();
        } catch (verifyErr: any) {
            const msg = verifyErr.message || '';
            if (msg.includes('535') || msg.includes('Username and Password not accepted') || msg.includes('BadCredentials')) {
                throw new AppError(
                    'Gmail rejected this App Password (535 Bad Credentials). Please ensure 2-Step Verification is active on your Google Account and generate a 16-character App Password at https://myaccount.google.com/apppasswords.',
                    400
                );
            }
            throw new AppError(`Gmail SMTP verification failed: ${msg}`, 400);
        }

        // Encrypt app password
        const appPasswordEncrypted = encrypt(cleanPassword);

        // Enforce single sender account: delete any existing credentials for this user
        await prisma.emailCredential.deleteMany({
            where: { userId },
        });

        // Create the single verified sender credential
        const credential = await prisma.emailCredential.create({
            data: {
                userId,
                email: cleanEmail,
                appPasswordEncrypted,
                isDefault: true,
            },
            select: {
                id: true,
                email: true,
                isDefault: true,
                createdAt: true,
            },
        });

        return credential;
    }

    async getCredentials(userId: string) {
        const credentials = await prisma.emailCredential.findMany({
            where: { userId },
            select: {
                id: true,
                email: true,
                isDefault: true,
                createdAt: true,
            },
            orderBy: { createdAt: 'desc' },
        });

        return credentials;
    }

    async getCredentialById(userId: string, credentialId: string) {
        const credential = await prisma.emailCredential.findFirst({
            where: { id: credentialId, userId },
        });

        if (!credential) {
            throw new AppError('Credential not found', 404);
        }

        return {
            id: credential.id,
            email: credential.email,
            appPassword: decrypt(credential.appPasswordEncrypted),
            isDefault: credential.isDefault,
        };
    }

    async deleteCredential(userId: string, credentialId: string) {
        const credential = await prisma.emailCredential.findFirst({
            where: { id: credentialId, userId },
        });

        if (!credential) {
            throw new AppError('Credential not found', 404);
        }

        await prisma.emailCredential.delete({
            where: { id: credentialId },
        });

        return { message: 'Credential deleted successfully' };
    }
}

export default new CredentialService();
