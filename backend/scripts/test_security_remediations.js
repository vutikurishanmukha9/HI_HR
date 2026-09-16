const http = require('http');
const io = require('../../frontend/node_modules/socket.io-client');
const { generateAccessToken } = require('../dist/utils/jwt');
const prisma = require('../dist/config/database').default;

const API_PORT = process.env.PORT || 5000;
const BASE_URL = `http://localhost:${API_PORT}`;

function post(path, body = {}, headers = {}) {
    return new Promise((resolve, reject) => {
        const payload = JSON.stringify(body);
        const req = http.request(`${BASE_URL}${path}`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(payload),
                ...headers,
            },
        }, (res) => {
            let data = '';
            res.on('data', chunk => data += chunk);
            res.on('end', () => {
                let parsed;
                try {
                    parsed = JSON.parse(data);
                } catch {
                    parsed = data;
                }
                resolve({ statusCode: res.statusCode, data: parsed });
            });
        });
        req.on('error', reject);
        req.write(payload);
        req.end();
    });
}

async function runSecurityTests() {
    console.log('=== RUNNING SECURITY REMEDIATION VERIFICATION SUITE ===\n');

    // 1. Auth Refresh & Logout Test
    console.log('1. Testing Auth Refresh and Logout lifecycle...');
    const testUser = await prisma.user.findFirst({ where: { isActive: true } });
    if (!testUser) {
        throw new Error('No user found in dev.db to test auth');
    }

    const { generateRefreshToken } = require('../dist/utils/jwt');
    const validRefreshToken = generateRefreshToken({
        id: testUser.id,
        email: testUser.email,
        fullName: testUser.fullName,
    });

    // Test /api/auth/refresh
    const refreshRes = await post('/api/auth/refresh', { refreshToken: validRefreshToken });
    if (refreshRes.statusCode !== 200 || !refreshRes.data.accessToken) {
        throw new Error(`FAIL: Refresh token endpoint failed: ${JSON.stringify(refreshRes)}`);
    }
    console.log('PASS: /api/auth/refresh successfully issued new access token');

    // Test /api/auth/logout
    const logoutRes = await post('/api/auth/logout');
    if (logoutRes.statusCode !== 200) {
        throw new Error(`FAIL: Logout endpoint failed: ${JSON.stringify(logoutRes)}`);
    }
    console.log('PASS: /api/auth/logout responded with status 200\n');

    // 2. Socket.IO Auth Enforcement
    console.log('2. Testing Socket.IO unauthorized connection rejection...');
    const unauthenticatedSocket = io(BASE_URL, {
        transports: ['websocket'],
        reconnection: false,
    });

    const unauthRejected = await new Promise((resolve) => {
        unauthenticatedSocket.on('connect_error', (err) => {
            console.log('Unauthenticated socket rejected with error:', err.message);
            unauthenticatedSocket.disconnect();
            resolve(true);
        });
        unauthenticatedSocket.on('connect', () => {
            unauthenticatedSocket.disconnect();
            resolve(false);
        });
    });

    if (!unauthRejected) {
        throw new Error('FAIL: Socket.IO allowed unauthenticated connection!');
    }
    console.log('PASS: Socket.IO rejected unauthenticated connection cleanly');

    console.log('\n3. Testing Socket.IO authenticated connection & room isolation...');
    const user1Token = generateAccessToken({
        id: testUser.id,
        email: testUser.email,
        fullName: testUser.fullName,
    });

    const authenticatedSocket = io(BASE_URL, {
        transports: ['websocket'],
        reconnection: false,
        auth: { token: user1Token },
    });

    const authConnected = await new Promise((resolve) => {
        authenticatedSocket.on('connect', () => resolve(true));
        authenticatedSocket.on('connect_error', (err) => {
            console.error('Authenticated socket failed to connect:', err.message);
            resolve(false);
        });
    });

    if (!authConnected) {
        throw new Error('FAIL: Authenticated socket failed to connect!');
    }
    console.log('PASS: Socket.IO connected with valid user JWT');

    // 4. Test unauthorized campaign join
    const campaignRunManager = require('../dist/services/campaignRun.manager').default;
    // Create run for someone else
    const victimRun = campaignRunManager.createRun('different-user-id-victim', ['test@example.com']);
    
    let receivedUnauthorizedError = false;
    authenticatedSocket.on('error', (err) => {
        if (err.message && err.message.includes('Unauthorized')) {
            receivedUnauthorizedError = true;
        }
    });

    authenticatedSocket.emit('join:campaign', victimRun.id);
    await new Promise(r => setTimeout(r, 600));

    if (!receivedUnauthorizedError) {
        throw new Error('FAIL: Socket was not blocked from joining victim run room!');
    }
    console.log('PASS: Socket was blocked with Unauthorized when joining unowned campaign run');
    authenticatedSocket.disconnect();

    // 5. Test dangerous attachment extension blocking
    console.log('\n4. Testing Attachment Security (dangerous extension blocking)...');
    const campaignPayload = {
        credentialEmail: testUser.email,
        subject: 'Job application',
        body: 'Hello HR',
        recipients: [{ email: 'hr@example.com' }],
        attachments: [
            {
                filename: 'malicious_resume.exe',
                contentType: 'application/octet-stream',
                content: Buffer.from('fake exe content').toString('base64'),
            }
        ]
    };

    const blockRes = await post('/api/campaigns/send', campaignPayload, {
        'Authorization': `Bearer ${user1Token}`,
    });

    if (blockRes.statusCode !== 400 || !JSON.stringify(blockRes.data).includes('disallowed file extension')) {
        throw new Error(`FAIL: Dangerous attachment was not rejected with 400! Got ${blockRes.statusCode}: ${JSON.stringify(blockRes.data)}`);
    }
    console.log('PASS: Dangerous executable attachment (.exe) was blocked with 400 INVALID_FILE_TYPE');

    console.log('\n===========================================================');
    console.log('ALL SECURITY REMEDIATION TESTS PASSED (5/5)!');
    console.log('===========================================================');
    process.exit(0);
}

runSecurityTests().catch(err => {
    console.error('TEST SUITE FAILED:', err);
    process.exit(1);
});
