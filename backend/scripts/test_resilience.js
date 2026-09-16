const fs = require('fs');
const path = require('path');

const TRACKING_FILE = path.join(process.cwd(), 'logs', 'email_tracking.jsonl');

async function testResilience() {
    console.log('=== TESTING TRACKING FILE CORRUPTION RESILIENCE ===\n');

    const trackingService = require('../dist/services/tracking.service').default;

    // 1. Create a valid record
    const { trackingToken } = await trackingService.createTrackingRecord({
        recipientEmail: 'good_record@example.com',
        subject: 'Good record',
        campaignId: 'resilience-test'
    });
    console.log('Created valid record with token:', trackingToken);

    // 2. Inject corrupted/malformed lines into the JSONL file
    console.log('Injecting corrupted lines into email_tracking.jsonl...');
    fs.appendFileSync(TRACKING_FILE, '{"bad_json": true, incomplete...\n');
    fs.appendFileSync(TRACKING_FILE, '\n\n   \n');
    fs.appendFileSync(TRACKING_FILE, 'RANDOM_GARBAGE_LINE_NOT_JSON\n');

    // 3. Create another valid record
    const { trackingToken: token2 } = await trackingService.createTrackingRecord({
        recipientEmail: 'second_good@example.com',
        subject: 'Second good',
        campaignId: 'resilience-test'
    });
    console.log('Created second valid record with token:', token2);

    // 4. Query stats
    const stats = trackingService.getTrackingStats('resilience-test');
    console.log('Total records recovered:', stats.totalSent);

    if (stats.totalSent !== 2) {
        throw new Error(`FAIL: Expected 2 valid records, got ${stats.totalSent}!`);
    }

    console.log('PASS: System successfully parsed valid records and skipped corrupted lines without crashing or losing data!\n');

    // 5. Clean up
    trackingService.clearTrackingRecords('resilience-test');
    console.log('Cleaned up test records.');
}

testResilience().catch(err => {
    console.error('RESILIENCE TEST FAILED:', err);
    process.exit(1);
});
