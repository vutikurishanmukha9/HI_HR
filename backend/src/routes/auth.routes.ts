import { Router } from 'express';
import authController from '../controllers/auth.controller';
import { validate } from '../middleware/validation';
import { authenticate } from '../middleware/auth';
import { authLimiter } from '../middleware/rateLimit';
import { registerSchema, loginSchema } from '../utils/validation';

const router = Router();

// Public routes with rate limiting
router.post('/register', authLimiter, validate(registerSchema), authController.register);
router.post('/login', authLimiter, validate(loginSchema), authController.login);

// Protected routes
router.get('/me', authenticate, authController.getMe);

export default router;
