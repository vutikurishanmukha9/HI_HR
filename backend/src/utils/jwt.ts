const jwt = require('jsonwebtoken');
import { env } from '../config/env';

interface TokenPayload {
    id: string;
    email: string;
    fullName: string;
}

export const generateAccessToken = (payload: TokenPayload) => {
    return jwt.sign(payload, env.JWT_SECRET, {
        expiresIn: env.JWT_EXPIRES_IN,
    });
};

export const generateRefreshToken = (payload: TokenPayload) => {
    return jwt.sign(payload, env.REFRESH_TOKEN_SECRET, {
        expiresIn: env.REFRESH_TOKEN_EXPIRES_IN,
    });
};

export const verifyAccessToken = (token: string) => {
    try {
        return jwt.verify(token, env.JWT_SECRET) as TokenPayload;
    } catch (error) {
        throw new Error('Invalid or expired token');
    }
};

export const verifyRefreshToken = (token: string) => {
    try {
        return jwt.verify(token, env.REFRESH_TOKEN_SECRET) as TokenPayload;
    } catch (error) {
        throw new Error('Invalid or expired refresh token');
    }
};
