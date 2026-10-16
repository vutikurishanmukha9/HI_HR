/**
 * Health & Diagnostics Routes
 */
import { Router } from 'express';
import { HealthController } from '../controllers/health.controller';

const router = Router();

// Liveness probe
router.get('/', HealthController.getLiveness);
router.get('/live', HealthController.getLiveness);

// Readiness probe
router.get('/ready', HealthController.getReadiness);

// System metrics
router.get('/metrics', HealthController.getMetrics);

export default router;
