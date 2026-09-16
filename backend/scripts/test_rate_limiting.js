const http = require('http');

function get(url) {
    return new Promise((resolve) => {
        http.get(url, (res) => {
            let data = '';
            res.on('data', chunk => data += chunk);
            res.on('end', () => resolve({ status: res.statusCode, data }));
        }).on('error', err => resolve({ status: 500, error: err.message }));
    });
}

async function testRateLimit() {
    console.log('Testing rate limiter on /api/track/open...');
    
    // We will make 105 rapid requests to a tracking token
    const token = 'test-token-rate-limit';
    let blockedCount = 0;
    let okCount = 0;

    for (let i = 1; i <= 105; i++) {
        const res = await get(`http://localhost:5000/api/track/open/${token}`);
        if (res.status === 429) {
            blockedCount++;
            if (blockedCount === 1) {
                console.log(`REQUEST #${i} WAS BLOCKED WITH 429! Body:`, res.data);
            }
        } else if (res.status === 200) {
            okCount++;
        }
    }

    console.log(`Results: ${okCount} OK, ${blockedCount} Rate-limited (429)`);
    if (blockedCount > 0) {
        console.log('CRITICAL DEFECT CONFIRMED: Tracking pixel IS blocked by apiLimiter!');
    } else {
        console.log('Tracking pixel is not blocked.');
    }
}

testRateLimit();
