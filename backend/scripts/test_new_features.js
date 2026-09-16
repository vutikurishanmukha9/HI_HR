const http = require('http');
const prisma = require('../dist/config/database').default;
const { generateAccessToken } = require('../dist/utils/jwt');

const API_PORT = process.env.PORT || 5000;
const BASE_URL = `http://localhost:${API_PORT}`;

function request(path, options = {}) {
    return new Promise((resolve, reject) => {
        const payload = options.body ? JSON.stringify(options.body) : null;
        const headers = {
            'Content-Type': 'application/json',
            ...(payload ? { 'Content-Length': Buffer.byteLength(payload) } : {}),
            ...(options.headers || {}),
        };

        const req = http.request(`${BASE_URL}${path}`, {
            method: options.method || 'GET',
            headers,
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
        if (payload) req.write(payload);
        req.end();
    });
}

async function runTests() {
    console.log('=== RUNNING BACKEND NEW FEATURES VERIFICATION SUITE ===\n');

    const testUser = await prisma.user.findFirst({ where: { isActive: true } });
    if (!testUser) throw new Error('No test user found in db');

    const token = generateAccessToken({
        id: testUser.id,
        email: testUser.email,
        fullName: testUser.fullName,
    });
    const authHeaders = { 'Authorization': `Bearer ${token}` };

    // 1. Test GET /api/templates
    console.log('1. Testing GET /api/templates (should return starter templates)...');
    const getRes = await request('/api/templates', { headers: authHeaders });
    if (getRes.statusCode !== 200 || !Array.isArray(getRes.data.templates)) {
        throw new Error(`FAIL: GET /api/templates failed: ${JSON.stringify(getRes)}`);
    }
    console.log(`Found ${getRes.data.templates.length} templates (including system templates)`);
    const systemTpl = getRes.data.templates.find(t => t.isSystem);
    if (!systemTpl) throw new Error('FAIL: System templates not found in listing!');
    console.log('PASS: System templates returned cleanly\n');

    // 2. Test POST /api/templates
    console.log('2. Testing POST /api/templates (creating custom template)...');
    const createRes = await request('/api/templates', {
        method: 'POST',
        headers: authHeaders,
        body: {
            name: 'My Custom Outreach Template',
            category: 'Tech Inquiries',
            subject: 'Excited about {companyName} engineering roles',
            body: '<p>Hello {fullName}, this is custom content.</p>'
        }
    });

    if (createRes.statusCode !== 201 || !createRes.data.template?.id) {
        throw new Error(`FAIL: POST /api/templates failed: ${JSON.stringify(createRes)}`);
    }
    const createdId = createRes.data.template.id;
    console.log('PASS: Custom template created with ID:', createdId, '\n');

    // 3. Test DELETE /api/templates/:id
    console.log('3. Testing DELETE /api/templates/:id...');
    const delRes = await request(`/api/templates/${createdId}`, {
        method: 'DELETE',
        headers: authHeaders
    });
    if (delRes.statusCode !== 200) {
        throw new Error(`FAIL: DELETE /api/templates failed: ${JSON.stringify(delRes)}`);
    }
    console.log('PASS: Custom template deleted successfully\n');

    // 4. Test DNS MX Domain Verification
    console.log('4. Testing Domain MX DNS Validation...');
    const mxService = require('../dist/services/emailValidator.service').default;
    const testEmails = [
        'recruiter@gmail.com',
        'engineer@microsoft.com',
        'nobody@thisdomaindoesnotexist9876543210.com'
    ];

    const mxResults = await mxService.validateEmailsMx(testEmails);
    console.log('MX Results:');
    for (const [email, res] of Object.entries(mxResults)) {
        console.log(` - ${email}: ${res.hasMx ? 'VALID (MX: ' + res.exchange + ')' : 'INVALID (' + res.error + ')'}`);
    }

    if (!mxResults['recruiter@gmail.com'].hasMx) throw new Error('FAIL: gmail.com should have valid MX!');
    if (!mxResults['engineer@microsoft.com'].hasMx) throw new Error('FAIL: microsoft.com should have valid MX!');
    if (mxResults['nobody@thisdomaindoesnotexist9876543210.com'].hasMx) throw new Error('FAIL: invalid domain should not have MX!');

    console.log('PASS: DNS MX pre-flight validation works accurately\n');

    // 5. Test POST /api/campaigns/validate-domains
    console.log('5. Testing POST /api/campaigns/validate-domains...');
    const apiMxRes = await request('/api/campaigns/validate-domains', {
        method: 'POST',
        headers: authHeaders,
        body: { emails: ['test@gmail.com', 'bad@nonexistentdomain12345.org'] }
    });

    if (apiMxRes.statusCode !== 200 || !apiMxRes.data.results) {
        throw new Error(`FAIL: /api/campaigns/validate-domains failed: ${JSON.stringify(apiMxRes)}`);
    }
    console.log('PASS: /api/campaigns/validate-domains API endpoint verified\n');

    console.log('===========================================================');
    console.log('ALL BACKEND SUITE TESTS PASSED (5/5)!');
    console.log('===========================================================');
    process.exit(0);
}

runTests().catch(err => {
    console.error('TEST FAILED:', err);
    process.exit(1);
});
