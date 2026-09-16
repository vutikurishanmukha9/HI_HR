const { PrismaClient } = require('@prisma/client');
const prisma = new PrismaClient();

async function check() {
    try {
        const users = await prisma.user.findMany();
        console.log('--- USERS (' + users.length + ') ---');
        for (const u of users) {
            console.log(`User: ${u.email} (id: ${u.id}, name: ${u.fullName})`);
        }

        const creds = await prisma.emailCredential.findMany();
        console.log('--- CREDENTIALS (' + creds.length + ') ---');
        for (const c of creds) {
            console.log(`Cred: ${c.email} (userId: ${c.userId}, default: ${c.isDefault})`);
        }

        const campaigns = await prisma.campaign.findMany();
        console.log('--- CAMPAIGNS (' + campaigns.length + ') ---');
        for (const camp of campaigns) {
            console.log(`Campaign: ${camp.name} (id: ${camp.id}, status: ${camp.status}, sent: ${camp.sentCount}/${camp.totalRecipients})`);
        }
    } catch (err) {
        console.error('Error during check:', err);
    } finally {
        await prisma.$disconnect();
    }
}

check();
