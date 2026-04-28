# StreamMail Backend API

Backend server for the StreamMail cold email outreach platform built with Node.js, Express, TypeScript, and PostgreSQL.

## Features

- 🔐 **JWT Authentication** - Secure user authentication with access and refresh tokens
- 🗄️ **PostgreSQL Database** - Robust data storage with Prisma ORM
- 🔒 **Encrypted Credentials** - AES-256 encryption for email app passwords
- 📧 **Email Campaign Management** - Create, schedule, and manage email campaigns
- 📊 **Recipient Management** - Upload recipients from Excel files
- 🛡️ **Security** - Rate limiting, CORS, helmet, input validation
- 📝 **TypeScript** - Full type safety throughout the codebase

## Prerequisites

- Node.js (v18 or higher)
- PostgreSQL database
- npm or yarn

## Installation

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Set up environment variables:**
   ```bash
   cp .env.example .env
   ```
   Then edit `.env` with your configuration:
   - `DATABASE_URL`: PostgreSQL connection string
   - `JWT_SECRET`: Secret key for JWT tokens
   - `ENCRYPTION_KEY`: 32-character key for AES-256 encryption
   - Other configuration as needed

3. **Set up the database:**
   ```bash
   npm run prisma:generate
   npm run prisma:migrate
   ```

4. **Start the development server:**
   ```bash
   npm run dev
   ```

The server will start on `http://localhost:5000`

## API Endpoints

### Authentication

- `POST /api/auth/register` - Register a new user
- `POST /api/auth/login` - Login user
- `GET /api/auth/me` - Get current user (requires auth)

### Email Credentials

- `POST /api/credentials` - Save email credentials (requires auth)
- `GET /api/credentials` - Get user's credentials (requires auth)
- `DELETE /api/credentials/:id` - Delete credential (requires auth)

### Campaigns

- `POST /api/campaigns` - Create new campaign (requires auth)
- `GET /api/campaigns` - Get user's campaigns (requires auth)
- `GET /api/campaigns/:id` - Get campaign details (requires auth)
- `PATCH /api/campaigns/:id` - Update campaign (requires auth)
- `DELETE /api/campaigns/:id` - Delete campaign (requires auth)
- `POST /api/campaigns/upload-recipients` - Upload recipients from Excel (requires auth)

## Project Structure

```
backend/
├── prisma/
│   └── schema.prisma          # Database schema
├── src/
│   ├── config/
│   │   ├── database.ts        # Prisma client
│   │   └── env.ts             # Environment config
│   ├── controllers/           # Request handlers
│   ├── middleware/            # Express middleware
│   ├── routes/                # API routes
│   ├── services/              # Business logic
│   ├── types/                 # TypeScript types
│   ├── utils/                 # Utility functions
│   └── server.ts              # Main server file
├── uploads/                   # File upload directory
├── .env.example               # Environment template
├── package.json
└── tsconfig.json
```

## Scripts

- `npm run dev` - Start development server with hot reload
- `npm run build` - Build for production
- `npm start` - Start production server
- `npm run prisma:generate` - Generate Prisma client
- `npm run prisma:migrate` - Run database migrations
- `npm run prisma:studio` - Open Prisma Studio

## Security Features

- **Password Hashing**: bcrypt with 10 salt rounds
- **JWT Tokens**: Short-lived access tokens (15min) + refresh tokens (7d)
- **Encryption**: AES-256-CBC for sensitive data
- **Rate Limiting**: 100 requests per 15 minutes
- **Input Validation**: Zod schemas for all endpoints
- **CORS**: Configured for frontend domain only
- **Helmet**: Security headers

## Database Schema

- **Users**: User accounts with authentication
- **EmailCredentials**: Encrypted email credentials
- **Campaigns**: Email campaigns with configuration
- **Recipients**: Campaign recipients with status tracking
- **Attachments**: File attachments for campaigns

## Development

The backend uses:
- **Express** for the web framework
- **Prisma** for database ORM
- **TypeScript** for type safety
- **Zod** for validation
- **Nodemailer** for email sending
- **JWT** for authentication

## Production Deployment

1. Build the project:
   ```bash
   npm run build
   ```

2. Set environment to production in `.env`:
   ```
   NODE_ENV=production
   ```

3. Run migrations:
   ```bash
   npm run prisma:migrate
   ```

4. Start the server:
   ```bash
   npm start
   ```

## License

Private - For demonstration purposes
