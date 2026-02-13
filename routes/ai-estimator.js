import express from 'express';
import multer from 'multer';
import path from 'path';
import fs from 'fs/promises';
import { z } from 'zod';
import { generateScope } from '../utils/openai-estimator.js';
import { generateExcel } from '../utils/excel.js';
import { sanitizeDivisions } from '../utils/sanitize.js';
import { requireAuth } from '../middleware/auth.js';

const router = express.Router();

// ─────────────────────────────────────────────
// Constants
// ─────────────────────────────────────────────
const MAX_PDF_MB = parseInt(process.env.MAX_UPLOAD_MB, 10) || 25;
const MAX_PDF_BYTES = MAX_PDF_MB * 1024 * 1024;
const ESTIMATOR_TMP_DIR = path.join(process.cwd(), 'tmp', 'estimator');

// Tighter JSON body limit for estimator routes (export payload)
router.use(express.json({ limit: '2mb' }));

// ─────────────────────────────────────────────
// Multer — PDF uploads only, disk-based, dedicated config
// ─────────────────────────────────────────────
const uploadPdf = multer({
  dest: ESTIMATOR_TMP_DIR,
  limits: { fileSize: MAX_PDF_BYTES },
  fileFilter: (_req, file, cb) => {
    if (file.mimetype === 'application/pdf') {
      cb(null, true);
    } else {
      cb(new multer.MulterError('LIMIT_UNEXPECTED_FILE', file.fieldname));
    }
  },
});

// ─────────────────────────────────────────────
// Validation schemas
// ─────────────────────────────────────────────
const generateBodySchema = z.object({
  projectName:    z.string().min(1, 'projectName is required'),
  projectAddress: z.string().default(''),
  version:        z.string().default('1.0'),
  date:           z.string().default(''),
});

const exportBodySchema = z.object({
  project: z.object({
    name:    z.string(),
    address: z.string().optional().default(''),
    version: z.string().optional().default('1.0'),
    date:    z.string().optional().default(''),
  }),
  divisions: z.array(z.object({
    division:  z.string(),
    items: z.array(z.object({
      title: z.string(),
      scope: z.string(),
    })),
  })),
});

// ─────────────────────────────────────────────
// Helper: delete temp file (silent on ENOENT)
// ─────────────────────────────────────────────
async function deleteTempFile(filePath) {
  if (!filePath) return;
  try {
    await fs.unlink(filePath);
    console.log(`  Cleaned up: ${path.basename(filePath)}`);
  } catch (err) {
    if (err.code !== 'ENOENT') {
      console.warn(`  Failed to clean up ${path.basename(filePath)}:`, err.message);
    }
  }
}

// ─────────────────────────────────────────────
// POST /api/ai-estimator/generate
// ─────────────────────────────────────────────
router.post('/generate', requireAuth, (req, res, next) => {
  // Wrap multer to catch file-type / size errors as JSON
  uploadPdf.single('file')(req, res, (err) => {
    if (err instanceof multer.MulterError) {
      if (err.code === 'LIMIT_FILE_SIZE') {
        return res.status(400).json({ error: 'PDF_TOO_LARGE', maxMB: MAX_PDF_MB });
      }
      if (err.code === 'LIMIT_UNEXPECTED_FILE') {
        return res.status(400).json({ error: 'INVALID_FILE_TYPE', message: 'Only PDF files are accepted' });
      }
      return res.status(400).json({ error: 'UPLOAD_ERROR', message: err.message });
    }
    if (err) {
      return res.status(400).json({ error: 'UPLOAD_ERROR', message: err.message });
    }
    next();
  });
}, async (req, res) => {
  let uploadedPath = null;
  let openaiFileId = null;

  try {
    // ── Validate file present ──
    if (!req.file) {
      return res.status(400).json({ error: 'No PDF file uploaded' });
    }
    uploadedPath = req.file.path;

    // ── Validate project metadata ──
    const projectData = generateBodySchema.parse(req.body);

    const sizeMB = (req.file.size / 1024 / 1024).toFixed(2);
    console.log(`\nAI Estimator - generate`);
    console.log(`  Project: ${projectData.projectName}`);
    console.log(`  File: ${req.file.originalname} (${sizeMB} MB)`);

    // ── Call OpenAI (streams PDF from disk, never loads full file into memory) ──
    const { result, fileId } = await generateScope(uploadedPath, projectData);
    openaiFileId = fileId;

    // ── Sanitize scope to strip any leaked prices/quantities ──
    const sanitized = {
      ...result,
      divisions: sanitizeDivisions(result.divisions || []),
    };

    console.log(`  Generated ${sanitized.divisions.length} division(s)`);

    return res.json(sanitized);

  } catch (error) {
    console.error('Generate failed:', error.message);

    if (error instanceof z.ZodError) {
      return res.status(400).json({ error: 'VALIDATION_ERROR', details: error.errors });
    }

    // OpenAI-specific errors
    if (error?.status || error?.code) {
      return res.status(502).json({
        error: 'OPENAI_FAILED',
        message: error.message?.slice(0, 200) || 'OpenAI request failed',
      });
    }

    return res.status(500).json({
      error: 'GENERATE_FAILED',
      message: error.message?.slice(0, 200) || 'Unknown error',
    });

  } finally {
    // CRITICAL: Always delete uploaded PDF — even on success
    await deleteTempFile(uploadedPath);
  }
});

// ─────────────────────────────────────────────
// POST /api/ai-estimator/export
// ─────────────────────────────────────────────
router.post('/export', requireAuth, async (req, res) => {
  try {
    const { project, divisions } = exportBodySchema.parse(req.body);

    console.log(`\nAI Estimator - export`);
    console.log(`  Project: ${project.name}`);

    const xlsxBuffer = await generateExcel(project, divisions);

    const filename = `ANDCON_SOW_${project.name.replace(/[^a-zA-Z0-9]/g, '_')}_v${project.version || '1'}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', xlsxBuffer.byteLength);

    console.log(`  Export complete: ${filename}`);

    return res.send(Buffer.from(xlsxBuffer));

  } catch (error) {
    console.error('Export failed:', error.message);

    if (error instanceof z.ZodError) {
      return res.status(400).json({ error: 'VALIDATION_ERROR', details: error.errors });
    }

    return res.status(500).json({
      error: 'EXPORT_FAILED',
      message: error.message?.slice(0, 200) || 'Unknown error',
    });
  }
});

export default router;
