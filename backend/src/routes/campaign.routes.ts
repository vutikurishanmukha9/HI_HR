import { Router } from 'express';
import campaignController, { upload } from '../controllers/campaign.controller';
import { authenticate } from '../middleware/auth';
import { validate, validateQuery } from '../middleware/validation';
import { createCampaignSchema, updateCampaignSchema, paginationSchema, sendCampaignSchema } from '../utils/validation';
import { emailLimiter, uploadLimiter } from '../middleware/rateLimit';

const router = Router();

// All routes require authentication
router.use(authenticate);

router.post('/', validate(createCampaignSchema), campaignController.createCampaign);
router.get('/', validateQuery(paginationSchema), campaignController.getCampaigns);
router.get('/:id', campaignController.getCampaignById);
router.patch('/:id', validate(updateCampaignSchema), campaignController.updateCampaign);
router.delete('/:id', campaignController.deleteCampaign);

// File upload route for recipients (with upload rate limit)
router.post('/upload-recipients', uploadLimiter, upload.single('file'), campaignController.uploadRecipients);

// Send campaign emails via backend (with email rate limit)
router.post('/send', emailLimiter, validate(sendCampaignSchema), campaignController.sendCampaign);

export default router;


