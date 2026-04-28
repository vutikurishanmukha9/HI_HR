import prisma from '../config/database';
import { encrypt, decrypt } from '../utils/encryption';
import { AppError } from '../middleware/errorHandler';
import { EmailCredentialDTO } from '../types';

export class CredentialService {
    async createCredential(userId: string, data: EmailCredentialDTO) {
        // Check if credential already exists
        const existing = await prisma.emailCredential.findUnique({
            where: {
                userId_email: {
                    userId,
                    email: data.email,
                },
            },
        });

        if (existing) {
            throw new AppError('Credential with this email already exists', 400);
        }

        // If this is set as default, unset other defaults
        if (data.isDefault) {
            await prisma.emailCredential.updateMany({
                where: { userId, isDefault: true },
                data: { isDefault: false },
            });
        }

        // Encrypt app password
        const appPasswordEncrypted = encrypt(data.appPassword);

        // Create credential
        const credential = await prisma.emailCredential.create({
            data: {
                userId,
                email: data.email,
                appPasswordEncrypted,
                isDefault: data.isDefault || false,
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
