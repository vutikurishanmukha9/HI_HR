# PostgreSQL Migration Guide

This document describes how to migrate from SQLite (development) to PostgreSQL (production).

## Prerequisites

- PostgreSQL 14+ installed
- Database created with appropriate user permissions

## Step 1: Update Environment Variables

```env
# Change from SQLite
DATABASE_URL="postgresql://user:password@localhost:5432/streammail?schema=public"
```

## Step 2: Update Prisma Schema

In `prisma/schema.prisma`, change the provider:

```prisma
datasource db {
  provider = "postgresql"
  url      = env("DATABASE_URL")
}
```

## Step 3: Generate Migration

```bash
# Create new migration
npx prisma migrate dev --name init_postgresql

# Or reset and create fresh migration
npx prisma migrate reset
npx prisma migrate dev --name init
```

## Step 4: Deploy to Production

```bash
# Apply migrations to production
npx prisma migrate deploy

# Generate Prisma Client
npx prisma generate
```

## Schema Differences

| SQLite | PostgreSQL |
|--------|------------|
| `TEXT` | `TEXT` or `VARCHAR` |
| `INTEGER` | `INTEGER` or `BIGINT` |
| `REAL` | `DOUBLE PRECISION` |
| No `ENUM` | Native `ENUM` support |

## Connection Pooling

For production, use connection pooling:

```env
DATABASE_URL="postgresql://user:password@localhost:5432/streammail?schema=public&connection_limit=5"
```

## Backup & Restore

```bash
# Backup
pg_dump streammail > backup.sql

# Restore
psql streammail < backup.sql
```
