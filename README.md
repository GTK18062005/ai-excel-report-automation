# AI Excel Report Platform

An AI-powered Excel report generation and analysis platform.

## Project Structure

```
ai-excel-report-platform/
├── backend/          # Spring Boot 3.3 + Java 21
├── frontend/         # Vite + React + TypeScript
└── docker-compose.yml
```

## Prerequisites

- Java 21+
- Node.js 20+
- Docker & Docker Compose

## Getting Started

### 1. Start PostgreSQL

```bash
docker-compose up -d
```

### 2. Start Backend

```bash
cd backend
$env:DB_PASSWORD="postgres"   # PowerShell
mvn spring-boot:run
```

The backend runs on **http://localhost:8080**

Health check: `GET /api/health` → `{"status": "UP"}`

### 3. Start Frontend

```bash
cd frontend
npm install
npm run dev
```

The frontend runs on **http://localhost:5173**

## Environment Variables

| Variable      | Description              | Default   |
|---------------|--------------------------|-----------|
| `DB_PASSWORD` | PostgreSQL password      | *(required)* |
