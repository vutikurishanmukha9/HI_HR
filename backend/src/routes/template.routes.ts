import { Router } from 'express';
import templateController from '../controllers/template.controller';
import { authenticate } from '../middleware/auth';

const router = Router();

// All template routes are protected by authentication
router.use(authenticate);

router.get('/', templateController.getTemplates);
router.post('/', templateController.createTemplate);
router.put('/:id', templateController.updateTemplate);
router.delete('/:id', templateController.deleteTemplate);

export default router;
