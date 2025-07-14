import express from 'express';
import cors from 'cors';
import fetch from 'isomorphic-fetch';
import mammoth from 'mammoth';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3001;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Env configs
const CONVEX_DEPLOYMENT_URL = process.env.CONVEX_DEPLOYMENT_URL;
const API_SECRET = process.env.FILE_PROCESSOR_SECRET;

// Helper: Chunk text
function chunkText(text, maxChunkSize = 1000) {
  if (!text?.trim()) return [];

  const sentences = text.split(/[.!?]+/).filter(s => s.trim());
  const chunks = [];
  let current = "";

  for (const s of sentences) {
    if (current.length + s.length + 1 > maxChunkSize) {
      if (current.trim()) chunks.push(current.trim());
      current = s.trim();
    } else {
      current += (current ? ". " : "") + s.trim();
    }
  }

  if (current.trim()) chunks.push(current.trim());
  return chunks;
}

// Extract content by file type
async function extractFileContent(fileBuffer, fileType, fileName) {
  try {
    let content = "";
    let metadata = { fileType, fileName };

    if (fileType.toLowerCase() === 'pdf') {
      const pdfParse = (await import('pdf-parse')).default;
      const data = await pdfParse(fileBuffer);
      content = data.text;
      metadata.pages = data.numpages;

    } else if (fileType.toLowerCase() === 'word') {
      const result = await mammoth.extractRawText({ buffer: fileBuffer });
      content = result.value;
      metadata.warnings = result.messages?.map(m => m.message);

    } else if (fileType.toLowerCase() === 'excel') {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(fileBuffer);
      const sheets = [];

      wb.eachSheet(ws => {
        const rows = [];
        ws.eachRow(row => {
          const line = row.values.slice(1).map(c => c?.toString()?.trim() || '').join('\t');
          if (line.trim()) rows.push(line);
        });
        if (rows.length) sheets.push(`Sheet: ${ws.name}\n${rows.join('\n')}`);
      });

      content = sheets.join('\n\n');
      metadata.sheets = wb.worksheets.length;

    } else if (fileType.toLowerCase() === 'powerpoint') {
      content = "PowerPoint content extraction is simplified.";
      metadata.note = "Basic placeholder. Use specialized service for deep extraction.";
    } else {
      throw new Error(`Unsupported file type: ${fileType}`);
    }

    return {
      content: content.trim(),
      metadata: {
        ...metadata,
        extractedAt: Date.now(),
        contentLength: content.length
      }
    };
  } catch (err) {
    console.error(`Extract error: ${err}`);
    throw new Error(`Failed to extract: ${err.message}`);
  }
}

// Convex API calls
async function getValidAccessToken(userEmail) {
  const res = await fetch(`${CONVEX_DEPLOYMENT_URL}/api/microsoftGraph/getValidAccessToken`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ userEmail }),
  });
  if (!res.ok) throw new Error("Invalid token response");
  const data = await res.json();
  return data.accessToken;
}

async function updateFileStatus(fileId, status, metadata = null, errorMessage = null) {
  await fetch(`${CONVEX_DEPLOYMENT_URL}/api/chatbotFilesMutations/updateFileStatus`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fileId, status, metadata, errorMessage }),
  });
}

async function insertChunkAndEmbedding(fileId, chunkIndex, text, embedding, createdAt) {
  const res = await fetch(`${CONVEX_DEPLOYMENT_URL}/api/chatbotFilesMutations/insertChunkAndEmbedding`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fileId, chunkIndex, text, embedding, createdAt }),
  });
  return res.ok;
}

// Download OneDrive file
async function downloadFileFromOneDrive(fileId, accessToken) {
  const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/content`, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) throw new Error("Failed to download from OneDrive");
  return await res.arrayBuffer();
}

// Embedding with OpenAI
async function createEmbedding(text) {
  const res = await fetch('https://api.openai.com/v1/embeddings', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${process.env.CONVEX_OPENAI_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      model: 'text-embedding-3-small',
      input: text,
      dimensions: 1536,
    }),
  });
  if (!res.ok) throw new Error("Embedding API failed");
  const result = await res.json();
  return result.data[0].embedding;
}

// Main endpoint
app.post('/process-file', async (req, res) => {
  const { fileId, oneDriveFileId, userEmail, fileName, fileType, secret } = req.body;

  if (secret !== API_SECRET) return res.status(401).json({ error: 'Unauthorized' });
  if (!fileId || !oneDriveFileId || !userEmail || !fileName || !fileType)
    return res.status(400).json({ error: 'Missing required parameters' });

  try {
    await updateFileStatus(fileId, 'Processing');

    const accessToken = await getValidAccessToken(userEmail);
    const fileBuffer = await downloadFileFromOneDrive(oneDriveFileId, accessToken);
    const { content, metadata } = await extractFileContent(fileBuffer, fileType, fileName);

    if (!content) throw new Error("Empty content extracted");

    const chunks = chunkText(content);
    if (!chunks.length) throw new Error("Chunking failed");

    const timestamp = Date.now();
    let successCount = 0;

    for (let i = 0; i < chunks.length; i++) {
      try {
        const embedding = await createEmbedding(chunks[i]);
        const ok = await insertChunkAndEmbedding(fileId, i, chunks[i], embedding, timestamp);
        if (ok) successCount++;
      } catch (e) {
        console.error(`Embedding failed at chunk ${i}:`, e.message);
      }
    }

    await updateFileStatus(fileId, 'Ready', {
      ...metadata,
      processedAt: Date.now(),
      embeddingsCount: successCount,
      chunksCreated: chunks.length,
    });

    res.json({ success: true, chunksCreated: chunks.length, embeddingsCreated: successCount });

  } catch (error) {
    console.error('Processing error:', error);
    await updateFileStatus(fileId, 'Failed', null, error.message);
    res.status(500).json({ success: false, error: error.message });
  }
});

// Health check
app.get('/health', (req, res) => {
  res.json({ status: 'OK', timestamp: new Date().toISOString() });
});

// Start server
app.listen(PORT, () => {
  console.log(`File processor listening at http://localhost:${PORT}/health`);
});
