const http = require('http');

function request(url, options = {}) {
    return new Promise((resolve, reject) => {
        const req = http.request(url, options, (res) => {
            let data = Buffer.alloc(0);
            res.on('data', chunk => data = Buffer.concat([data, chunk]));
            res.on('end', () => {
                resolve({
                    statusCode: res.statusCode,
                    headers: res.headers,
                    body: data,
                    text: data.toString('utf-8')
                });
            });
        });
        req.on('error', reject);
        req.end();
    });
}

async function testDelivery() {
    console.log('=== TESTING CLEAN DELIVERY LOGGING & MULTI-TENANT ISOLATION ===\n');

    const trackingService = require('../dist/services/tracking.service').default;

    const userId1 = 'user-tenant-1';
    const userId2 = 'user-tenant-2';
    const campaignId = 'test-clean-' + Date.now();

    // 1. Record 2 successful deliveries and 1 failure for userId1
    console.log('1. Recording sample deliveries for user 1...');
    await trackingService.recordDelivery({
        userId: userId1,
        recipientEmail: 'alex.hr@techcorp.com',
        subject: 'Job Inquiry: Senior Frontend Engineer',
        campaignId,
        status: 'sent'
    });

    await trackingService.recordDelivery({
        userId: userId1,
        recipientEmail: 'sarah.recruiter@globalinc.com',
        subject: 'Job Inquiry: Senior Frontend Engineer',
        campaignId,
        status: 'sent'
    });

    await trackingService.recordDelivery({
        userId: userId1,
        recipientEmail: 'invalid.mailbox@domaindoesnotexist.xyz',
        subject: 'Job Inquiry: Senior Frontend Engineer',
        campaignId,
        status: 'failed',
        errorMessage: 'Recipient email does not exist or was rejected by mail server'
    });

    // 2. Query stats for user 1 and user 2 (isolation check)
    console.log('2. Verifying multi-tenant isolation in stats...');
    const statsUser1 = trackingService.getTrackingStats(userId1, campaignId);
    const statsUser2 = trackingService.getTrackingStats(userId2, campaignId);

    console.log('User 1 Stats:', statsUser1);
    console.log('User 2 Stats:', statsUser2);

    if (statsUser1.totalSent !== 3 || statsUser1.sentCount !== 2 || statsUser1.failedCount !== 1 || statsUser1.successRate !== 67) {
        throw new Error('FAIL: User 1 stats did not match expected values!');
    }
    if (statsUser2.totalSent !== 0 || statsUser2.sentCount !== 0) {
        throw new Error('FAIL: Multi-tenant breach! User 2 can see User 1 stats!');
    }
    console.log('PASS: User isolation verified in delivery stats\n');

    // 3. Query paginated details
    console.log('3. Querying paginated delivery details for User 1...');
    const details = trackingService.getTrackingDetailsPaginated(userId1, 1, 10, campaignId);
    console.log('Total records for User 1:', details.records.length);
    details.records.forEach((r, idx) => {
        console.log(` [${idx + 1}] ${r.recipientEmail} | ${r.status.toUpperCase()} | ${r.errorMessage || 'Clean SMTP delivery'}`);
    });

    if (details.records.length !== 3) {
        throw new Error('FAIL: Expected 3 records for User 1!');
    }

    const detailsUser2 = trackingService.getTrackingDetailsPaginated(userId2, 1, 10, campaignId);
    if (detailsUser2.records.length !== 0) {
        throw new Error('FAIL: Multi-tenant breach! User 2 can see User 1 records!');
    }
    console.log('PASS: Paginated details isolated per-user cleanly\n');

    // 4. Verify no tracking pixel is generated
    console.log('4. Verifying email body has NO tracking pixel or redirect links...');
    const emailService = require('../dist/services/email.service').default;
    const testHtml = '<p>Check my portfolio at <a href="https://linkedin.com/in/myprofile">LinkedIn</a></p>';
    const prepared = await emailService.prepareTrackedEmail({
        recipientEmail: 'test@example.com',
        subject: 'Test Clean Email',
        html: testHtml,
        campaignId
    });

    console.log('Prepared HTML:', prepared.html);
    if (prepared.html.includes('/api/track/open') || prepared.html.includes('/api/track/click')) {
        throw new Error('FAIL: Prepared HTML still contains tracking URLs!');
    }
    if (!prepared.html.includes('href="https://linkedin.com/in/myprofile"')) {
        throw new Error('FAIL: Direct link was modified!');
    }
    console.log('PASS: Direct links preserved untouched, zero tracking pixels in email body\n');

    // 5. Clean up test records
    console.log('5. Resetting test delivery records for User 1...');
    trackingService.clearTrackingRecords(userId1, campaignId);
    const cleaned = trackingService.getTrackingStats(userId1, campaignId);
    if (cleaned.totalSent !== 0) {
        throw new Error('FAIL: Records were not cleared!');
    }
    console.log('PASS: Records cleared cleanly for User 1\n');

    console.log('===========================================================');
    console.log('ALL CLEAN DELIVERY & SECURITY ISOLATION TESTS PASSED (5/5)!');
    console.log('===========================================================');
}

testDelivery().catch(err => {
    console.error('TEST FAILED:', err);
    process.exit(1);
});
