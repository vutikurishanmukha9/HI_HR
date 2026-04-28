import { Response, NextFunction } from 'express';
import { AuthRequest } from '../types';
import credentialService from '../services/credential.service';

export class CredentialController {
    async createCredential(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const credential = await credentialService.createCredential(req.user!.id, req.body);
            res.status(201).json({ credential });
        } catch (error) {
            next(error);
        }
    }

    async getCredentials(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const credentials = await credentialService.getCredentials(req.user!.id);
            res.json({ credentials });
        } catch (error) {
            next(error);
        }
    }

    async deleteCredential(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const result = await credentialService.deleteCredential(req.user!.id, req.params.id);
            res.json(result);
        } catch (error) {
            next(error);
        }
    }
}

export default new CredentialController();
