import { Response, NextFunction } from 'express';
import { AuthRequest } from '../types';
import templateService from '../services/template.service';

export class TemplateController {
    async getTemplates(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const templates = await templateService.getTemplates(req.user!.id);
            res.json({ templates });
        } catch (error) {
            next(error);
        }
    }

    async createTemplate(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const template = await templateService.createTemplate(req.user!.id, req.body);
            res.status(201).json({ template });
        } catch (error) {
            next(error);
        }
    }

    async updateTemplate(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const template = await templateService.updateTemplate(req.user!.id, req.params.id, req.body);
            res.json({ template });
        } catch (error) {
            next(error);
        }
    }

    async deleteTemplate(req: AuthRequest, res: Response, next: NextFunction): Promise<void> {
        try {
            const result = await templateService.deleteTemplate(req.user!.id, req.params.id);
            res.json(result);
        } catch (error) {
            next(error);
        }
    }
}

export default new TemplateController();
