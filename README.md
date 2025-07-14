# File Processor Service

External service for processing chatbot files (PDF, Word, Excel, PowerPoint) and creating embeddings.

## Setup

1. Install dependencies:
```bash
cd file-processor-service
npm install
```

2. Copy environment variables:
```bash
cp .env.example .env
```

3. Configure environment variables in `.env`:
- `CONVEX_DEPLOYMENT_URL`: Your Convex deployment URL
- `CONVEX_OPENAI_API_KEY`: OpenAI API key for embeddings
- `FILE_PROCESSOR_SECRET`: Secret key for API authentication
- `PORT`: Port for the service (default: 3001)

4. Start the service:
```bash
npm start
```

## API Endpoints

### POST /process-file
Process a file and create embeddings.

**Request Body:**
```json
{
  "fileId": "convex-file-id",
  "oneDriveFileId": "onedrive-file-id",
  "userEmail": "user@example.com",
  "fileName": "document.pdf",
  "fileType": "PDF",
  "secret": "your-secret-key"
}
```

**Response:**
```json
{
  "success": true,
  "message": "File processed successfully",
  "chunksCreated": 15,
  "embeddingsCreated": 15,
  "contentLength": 12450
}
```

### GET /health
Health check endpoint.

## Supported File Types

- **PDF**: Uses `pdf-parse` library
- **Word (.docx)**: Uses `mammoth` library
- **Excel (.xlsx)**: Uses `exceljs` library
- **PowerPoint (.pptx)**: Basic text extraction

## Deployment

This service can be deployed to:
- Railway
- Vercel (as serverless functions)
- Heroku
- Any Node.js hosting platform

Make sure to set the environment variables in your deployment platform.
