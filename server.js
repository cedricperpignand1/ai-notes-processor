import express from 'express';
import multer from 'multer';
import cors from 'cors';
import dotenv from 'dotenv';
import path from 'path';
import fs from 'fs/promises';
import { z } from 'zod';
import { extractAudio, checkFFmpeg } from './utils/ffmpeg.js';
import { transcribeAudio, generateNotes } from './utils/openai.js';
import { sendDailyLogEmail, verifyEmailConfig } from './utils/email.js';
import { cleanupTempFiles, generateJobId } from './utils/cleanup.js';
import { requireAuth } from './middleware/auth.js';
import aiEstimatorRouter from './routes/ai-estimator.js';

// Load environment variables
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3001;

// CORS — restrict to allowed origins in production
const allowedOrigins = process.env.ALLOWED_ORIGINS
  ? process.env.ALLOWED_ORIGINS.split(',').map(o => o.trim())
  : ['http://localhost:3000', 'http://localhost:3001'];

app.use(cors({
  origin: (origin, callback) => {
    // Allow requests with no origin (server-to-server, curl)
    if (!origin) return callback(null, true);
    if (allowedOrigins.includes(origin)) return callback(null, true);
    return callback(new Error(`CORS: origin ${origin} not allowed`));
  },
  credentials: true,
}));
app.use(express.json({ limit: '10mb' }));

// Ensure tmp directory exists
const TMP_DIR = path.join(process.cwd(), 'tmp');
await fs.mkdir(TMP_DIR, { recursive: true });

// Multer configuration for file uploads (300MB limit)
const upload = multer({
  dest: TMP_DIR,
  limits: {
    fileSize: 300 * 1024 * 1024, // 300MB
  },
  fileFilter: (req, file, cb) => {
    const allowedMimes = ['video/mp4', 'video/quicktime', 'video/x-msvideo'];
    if (allowedMimes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error('Only video files (.mp4, .mov) are allowed'));
    }
  },
});

// Validation schemas
const processSchema = z.object({
  projectName: z.string().min(1, 'Project name is required'),
  createdBy: z.string().min(1, 'Created by is required'),
});

const sendSchema = z.object({
  projectName: z.string().min(1),
  createdBy: z.string().min(1),
  notesText: z.string().min(1, 'Notes text is required'),
});

// Auth middleware imported from ./middleware/auth.js

// ── AI Estimator routes ──────────────────────────────────────
app.use('/api/ai-estimator', aiEstimatorRouter);

// Health check endpoint (no auth required)
app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    service: 'ai-notes-processor',
    timestamp: new Date().toISOString()
  });
});

// POST /api/ai-notes/process
// Upload video, extract audio, transcribe, generate notes
app.post('/api/ai-notes/process', requireAuth, upload.single('videoFile'), async (req, res) => {
  const jobId = generateJobId();
  let videoPath = null;
  let audioPath = null;

  try {
    // Validate request body
    const { projectName, createdBy } = processSchema.parse(req.body);

    // Check if video file was uploaded
    if (!req.file) {
      return res.status(400).json({ error: 'No video file uploaded' });
    }

    console.log(`\n🎬 Processing job ${jobId}`);
    console.log(`   Project: ${projectName}`);
    console.log(`   Created By: ${createdBy}`);
    console.log(`   File: ${req.file.originalname} (${(req.file.size / 1024 / 1024).toFixed(2)} MB)`);

    // Rename uploaded file to jobId.mp4
    videoPath = path.join(TMP_DIR, `${jobId}.mp4`);
    await fs.rename(req.file.path, videoPath);

    // Extract audio
    audioPath = path.join(TMP_DIR, `${jobId}.mp3`);
    await extractAudio(videoPath, audioPath);

    // Transcribe audio
    const transcriptText = await transcribeAudio(audioPath);

    // Generate notes
    const notesText = await generateNotes(transcriptText, projectName, createdBy);

    console.log(`✓ Job ${jobId} completed successfully\n`);

    // Return response
    res.json({
      jobId,
      transcriptText,
      notesText,
    });

  } catch (error) {
    console.error(`❌ Job ${jobId} failed:`, error.message);

    // Handle validation errors
    if (error instanceof z.ZodError) {
      return res.status(400).json({
        error: 'Validation error',
        details: error.errors
      });
    }

    // Handle multer errors
    if (error instanceof multer.MulterError) {
      if (error.code === 'LIMIT_FILE_SIZE') {
        return res.status(400).json({ error: 'File too large (max 300MB)' });
      }
      return res.status(400).json({ error: error.message });
    }

    // Generic error
    res.status(500).json({
      error: 'Processing failed',
      message: error.message
    });

  } finally {
    // CRITICAL: Always cleanup temp files, even if processing failed
    await cleanupTempFiles(jobId);
  }
});

// POST /api/ai-notes/send
// Send daily log email via Gmail SMTP
app.post('/api/ai-notes/send', requireAuth, async (req, res) => {
  try {
    // Validate request body
    const { projectName, createdBy, notesText } = sendSchema.parse(req.body);

    console.log(`\n📧 Sending email for project: ${projectName}`);

    // Send email
    await sendDailyLogEmail({ projectName, createdBy, notesText });

    console.log(`✓ Email sent successfully\n`);

    res.json({ ok: true });

  } catch (error) {
    console.error('❌ Email send failed:', error.message);

    if (error instanceof z.ZodError) {
      return res.status(400).json({
        error: 'Validation error',
        details: error.errors
      });
    }

    res.status(500).json({
      error: 'Email sending failed',
      message: error.message
    });
  }
});

// 404 handler
app.use((req, res) => {
  res.status(404).json({ error: 'Not found' });
});

// Error handler
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({
    error: 'Internal server error',
    message: err.message
  });
});

// Startup checks
async function startupChecks() {
  console.log('🚀 Starting AI Notes Processor Service\n');

  // Check required environment variables
  const requiredEnvVars = [
    'OPENAI_API_KEY',
    'GMAIL_USER',
    'GMAIL_APP_PASSWORD',
    'EMAIL_TO',
    'PROCESSOR_SHARED_SECRET',
  ];

  const missingVars = requiredEnvVars.filter(v => !process.env[v]);
  if (missingVars.length > 0) {
    console.error('❌ Missing required environment variables:');
    missingVars.forEach(v => console.error(`   - ${v}`));
    console.error('\nPlease create a .env file with all required variables.');
    console.error('See .env.example for reference.\n');
    process.exit(1);
  }

  // Check ffmpeg installation
  const hasFFmpeg = await checkFFmpeg();
  if (!hasFFmpeg) {
    console.error('❌ FFmpeg not found!');
    console.error('Please install FFmpeg:');
    console.error('  Windows: choco install ffmpeg  OR  download from https://ffmpeg.org/');
    console.error('  macOS:   brew install ffmpeg');
    console.error('  Linux:   apt install ffmpeg\n');
    process.exit(1);
  }
  console.log('✓ FFmpeg installed');

  // Verify email configuration
  const emailOk = await verifyEmailConfig();
  if (!emailOk) {
    console.error('❌ Gmail SMTP configuration failed!');
    console.error('Please check your GMAIL_USER and GMAIL_APP_PASSWORD.');
    console.error('Note: You need a Gmail App Password, not your regular password.\n');
    process.exit(1);
  }
  console.log('✓ Gmail SMTP configured');

  console.log(`✓ Temp directory: ${TMP_DIR}`);
  console.log(`✓ All checks passed\n`);
}

// Start server
await startupChecks();

app.listen(PORT, () => {
  console.log(`🎬 Processor service running on http://localhost:${PORT}`);
  console.log(`📧 Email recipients: ${process.env.EMAIL_TO}`);
  console.log(`\nEndpoints:`);
  console.log(`   GET  /health`);
  console.log(`   POST /api/ai-notes/process`);
  console.log(`   POST /api/ai-notes/send`);
  console.log(`   POST /api/ai-estimator/generate`);
  console.log(`   POST /api/ai-estimator/export\n`);
  console.log(`Press Ctrl+C to stop\n`);
});
