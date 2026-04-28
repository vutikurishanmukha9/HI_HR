import { Router } from 'express';
import credentialController from '../controllers/credential.controller';
import { authenticate } from '../middleware/auth';
import { validate } from '../middleware/validation';
import { emailCredentialSchema } from '../utils/validation';

const router = Router();

// All routes require authentication
router.use(authenticate);

router.post('/', validate(emailCredentialSchema), credentialController.createCredential);
router.get('/', credentialController.getCredentials);
router.delete('/:id', credentialController.deleteCredential);

export default router;
