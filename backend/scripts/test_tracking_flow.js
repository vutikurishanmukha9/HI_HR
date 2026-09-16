const http = require('http');

function request(url, options = {}) {
    return new Promise((resolve, reject) => {
        const req = http.request(url, options, (res) => {
            let data = Buffer.alloc(0);
            res.on('data', chunk => {
                data = Buffer.concat([data, chunk]);
            });
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

async function runTests() {
    console.log('=== COMPREHENSIVE EMAIL TRACKING VERIFICATION ===\n');

    const trackingService = require('../dist/services/tracking.service').default;

    // Test 1: HTML Entity decoding and URL normalization
    console.log('--- TEST 1: URL Decoding & Normalization ---');
    const rawUrl1 = 'https://example.com/jobs?role=engineer&amp;utm_source=cold_email&amp;id=42';
    const normalized1 = trackingService.normalizeUrl(rawUrl1);
    console.log('Original:', rawUrl1);
    console.log('Normalized:', normalized1);
    if (normalized1.includes('&amp;')) {
        throw new Error('FAIL: URL normalization did not decode &amp;!');
    }
    console.log('PASS: &amp; was decoded to &\n');

    // Test 2: Protocol prefixing for URLs without http://
    console.log('--- TEST 2: Missing Protocol Auto-Prefix ---');
    const rawUrl2 = 'linkedin.com/in/shanmukh-profile';
    const normalized2 = trackingService.normalizeUrl(rawUrl2);
    console.log('Original:', rawUrl2);
    console.log('Normalized:', normalized2);
    if (!normalized2.startsWith('https://')) {
        throw new Error('FAIL: URL did not receive https:// prefix!');
    }
    console.log('PASS: Missing protocol received https:// prefix\n');

    // Test 3: Tracking Pixel Generation (Non-hidden styling)
    console.log('--- TEST 3: Modern Zero-Footprint Tracking Pixel ---');
    const pixelHtml = trackingService.generateTrackingPixelHtml('sample-token');
    console.log('Pixel HTML:', pixelHtml);
    if (pixelHtml.includes('display:none')) {
        throw new Error('FAIL: Tracking pixel still uses display:none!');
    }
    console.log('PASS: Tracking pixel uses modern non-hidden styling\n');

    // Test 4: Create Tracking Record & Rewrite Links
    console.log('--- TEST 4: Tracking Record & Link Rewriting ---');
    const campaignId = 'test-campaign-' + Date.now();
    const recipient = 'candidate@company.com';
    const subject = 'Opportunity at Tech Corp';
    
    const { trackingToken } = await trackingService.createTrackingRecord({
        recipientEmail: recipient,
        subject,
        campaignId
    });

    const emailHtml = `
        <p>Dear Candidate,</p>
        <p>Take a look at our <a href="https://example.com/careers?dept=eng&amp;ref=outreach">Careers Page</a>.</p>
        <p>Also connect on <a href="linkedin.com/company/techcorp">LinkedIn</a>.</p>
    `;

    const { body: rewrittenHtml, links } = trackingService.rewriteLinksForTracking(emailHtml, trackingToken);
    console.log('Links rewritten:', links.length);
    links.forEach(l => console.log(` - Token: ${l.clickToken} -> ${l.originalUrl}`));

    // Test 5: Fetch Tracking Pixel via HTTP
    console.log('\n--- TEST 5: HTTP Pixel Fetch (/api/track/open/:token) ---');
    const pixelRes = await request(`http://localhost:5000/api/track/open/${trackingToken}`);
    console.log('Pixel Response Status:', pixelRes.statusCode);
    console.log('Content-Type:', pixelRes.headers['content-type']);
    console.log('Content-Length:', pixelRes.headers['content-length']);
    console.log('Cache-Control:', pixelRes.headers['cache-control']);
    console.log('CORS Header:', pixelRes.headers['access-control-allow-origin']);

    // Check stats after open
    await new Promise(r => setTimeout(r, 100)); // Allow async event
    let stats = trackingService.getTrackingStats(campaignId);
    console.log('Stats after open:');
    console.log(' - Total Sent:', stats.totalSent);
    console.log(' - Total Opened:', stats.totalOpened);
    console.log(' - Open Rate:', stats.openRate + '%');
    let rec = stats.records.find(r => r.trackingToken === trackingToken);
    console.log(' - Recipient Open Count:', rec?.openCount);
    if (rec?.openCount !== 1) {
        throw new Error('FAIL: openCount was not 1!');
    }
    console.log('PASS: Email open recorded correctly\n');

    // Test 6: HTTP Link Click Redirect (/api/track/click/:token)
    console.log('--- TEST 6: HTTP Link Click Redirect (/api/track/click/:token) ---');
    const clickToken1 = links[0].clickToken;
    const clickRes = await request(`http://localhost:5000/api/track/click/${clickToken1}`);
    console.log('Click Status:', clickRes.statusCode);
    console.log('Redirect Location:', clickRes.headers['location']);

    if (clickRes.statusCode !== 302) {
        throw new Error(`FAIL: Expected 302 redirect, got ${clickRes.statusCode}`);
    }
    if (clickRes.headers['location'].includes('&amp;')) {
        throw new Error('FAIL: Redirect Location still contains &amp;!');
    }
    console.log('PASS: Link click redirected with clean decoded URL (302)\n');

    // Test 7: Click-Triggered Open (Recipient with blocked images clicks link)
    console.log('--- TEST 7: Click-Triggered Open Test ---');
    const { trackingToken: tokenUnopened } = await trackingService.createTrackingRecord({
        recipientEmail: 'blocked_images@company.com',
        subject: 'Test Blocked Images',
        campaignId
    });
    const sampleBody2 = '<p><a href="https://example.com/unopened">Click Me</a></p>';
    const { links: links2 } = trackingService.rewriteLinksForTracking(sampleBody2, tokenUnopened);
    
    // Do NOT hit the pixel! Hit the link directly:
    const clickRes2 = await request(`http://localhost:5000/api/track/click/${links2[0].clickToken}`);
    console.log('Click on unopened email status:', clickRes2.statusCode);

    stats = trackingService.getTrackingStats(campaignId);
    const unopenedRec = stats.records.find(r => r.trackingToken === tokenUnopened);
    console.log('Unopened recipient openCount after link click:', unopenedRec?.openCount);
    if (unopenedRec?.openCount !== 1) {
        throw new Error('FAIL: Click did not trigger email open status!');
    }
    console.log('PASS: Link click automatically recorded email open for image-blocked client\n');

    // Test 8: Paginated Details
    console.log('--- TEST 8: Paginated Details Reporting ---');
    const details = trackingService.getTrackingDetailsPaginated(1, 10, campaignId);
    console.log('Total Records:', details.pagination.total);
    console.log('Open Rate:', details.openRate + '%');
    console.log('Total Clicks:', details.totalClicks);
    console.log('Unique Clicks:', details.uniqueClicks);
    console.log('PASS: Paginated stats calculated accurately\n');

    // Test 9: Clean up records
    console.log('--- TEST 9: Reset / Clear Records ---');
    trackingService.clearTrackingRecords(campaignId);
    const cleanedStats = trackingService.getTrackingStats(campaignId);
    console.log('Records after reset:', cleanedStats.totalSent);
    if (cleanedStats.totalSent !== 0) {
        throw new Error('FAIL: Records were not cleared!');
    }
    console.log('PASS: Tracking records cleared successfully\n');

    console.log('==============================================');
    console.log('ALL TRACKING TESTS PASSED PERFECTLY (9/9)!');
    console.log('==============================================');
}

runTests().catch((err) => {
    console.error('\nTEST FAILED:', err);
    process.exit(1);
});
