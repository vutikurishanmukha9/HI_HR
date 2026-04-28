import bcrypt from 'bcryptjs';
import prisma from '../config/database';
import { generateAccessToken, generateRefreshToken } from '../utils/jwt';
import { AppError } from '../middleware/errorHandler';
import { RegisterDTO, LoginDTO } from '../types';

export class AuthService {
    async register(data: RegisterDTO) {
        // Check if user already exists
        const existingUser = await prisma.user.findUnique({
            where: { email: data.email },
        });

        if (existingUser) {
            throw new AppError('User with this email already exists', 400);
        }

        // Hash password
        const passwordHash = await bcrypt.hash(data.password, 10);

        // Create user
        const user = await prisma.user.create({
            data: {
                email: data.email,
                passwordHash,
                fullName: data.fullName,
            },
            select: {
                id: true,
                email: true,
                fullName: true,
                createdAt: true,
            },
        });

        // Generate tokens
        const accessToken = generateAccessToken({
            id: user.id,
            email: user.email,
            fullName: user.fullName,
        });

        const refreshToken = generateRefreshToken({
            id: user.id,
            email: user.email,
            fullName: user.fullName,
        });

        return {
            user,
            accessToken,
            refreshToken,
        };
    }

    async login(data: LoginDTO) {
        // Find user
        const user = await prisma.user.findUnique({
            where: { email: data.email },
        });

        if (!user || !user.isActive) {
            throw new AppError('Invalid credentials', 401);
        }

        // Check if account is locked
        if (user.lockedUntil && user.lockedUntil > new Date()) {
            const remainingMinutes = Math.ceil((user.lockedUntil.getTime() - Date.now()) / 60000);
            throw new AppError(`Account is locked. Try again in ${remainingMinutes} minutes.`, 423);
        }

        // Verify password
        const isPasswordValid = await bcrypt.compare(data.password, user.passwordHash);

        if (!isPasswordValid) {
            // Increment failed attempts
            const newFailedAttempts = user.failedLoginAttempts + 1;
            const updateData: { failedLoginAttempts: number; lockedUntil?: Date } = {
                failedLoginAttempts: newFailedAttempts,
            };

            // Lock account after 5 failed attempts
            if (newFailedAttempts >= 5) {
                updateData.lockedUntil = new Date(Date.now() + 30 * 60 * 1000); // 30 minutes
            }

            await prisma.user.update({
                where: { id: user.id },
                data: updateData,
            });

            const remainingAttempts = Math.max(0, 5 - newFailedAttempts);
            if (remainingAttempts > 0) {
                throw new AppError(`Invalid credentials. ${remainingAttempts} attempts remaining.`, 401);
            } else {
                throw new AppError('Account locked for 30 minutes due to too many failed attempts.', 423);
            }
        }

        // Reset failed attempts on successful login
        await prisma.user.update({
            where: { id: user.id },
            data: {
                lastLogin: new Date(),
                failedLoginAttempts: 0,
                lockedUntil: null,
            },
        });

        // Generate tokens
        const accessToken = generateAccessToken({
            id: user.id,
            email: user.email,
            fullName: user.fullName,
        });

        const refreshToken = generateRefreshToken({
            id: user.id,
            email: user.email,
            fullName: user.fullName,
        });

        return {
            user: {
                id: user.id,
                email: user.email,
                fullName: user.fullName,
                createdAt: user.createdAt,
            },
            accessToken,
            refreshToken,
        };
    }

    async getMe(userId: string) {
        const user = await prisma.user.findUnique({
            where: { id: userId },
            select: {
                id: true,
                email: true,
                fullName: true,
                createdAt: true,
                lastLogin: true,
            },
        });

        if (!user) {
            throw new AppError('User not found', 404);
        }

        return user;
    }
}

export default new AuthService();
