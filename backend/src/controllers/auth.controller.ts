import { Request, Response, NextFunction } from 'express';
import { AuthRequest } from '../types';
import authService from '../services/auth.service';

export class AuthController {
    async register(req: Request, res: Response, next: NextFunction): Promise<void> {
        try {
            const result = await authService.register(req.body);
            res.status(201).json(result);
        } catch (error) {
            next(error);
        }
    }

    async login(req: Request, res: Response, next: NextFunction): Promise<void> {
        try {
            const result = await authService.login(req.body);
            res.json(result);
        } catch (error) {
            next(error);
        }
    }

    async getMe(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const user = await authService.getMe(req.user!.id);
            res.json({ user });
        } catch (error) {
            next(error);
        }
    }
}

export default new AuthController();
