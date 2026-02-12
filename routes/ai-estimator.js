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
// Multer — PDF uploads only, 50MB max
// ─────────────────────────────────────────────
const TMP_DIR = path.join(process.cwd(), 'tmp');

const upload = multer({
  dest: TMP_DIR,
  limits: {
    fileSize: (parseInt(process.env.MAX_UPLOAD_MB, 10) || 50) * 1024 * 1024,
  },
  fileFilter: (_req, file, cb) => {
    if (file.mimetype === 'application/pdf') {
      cb(null, true);
    } else {
      cb(new Error('Only PDF files are accepted'));
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
// POST /api/ai-estimator/generate
// ─────────────────────────────────────────────
router.post('/generate', requireAuth, upload.single('file'), async (req, res) => {
  let uploadedPath = null;

  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No PDF file uploaded' });
    }

    uploadedPath = req.file.path;

    // Validate project metadata from form fields
    const projectData = generateBodySchema.parse(req.body);

    console.log(`\n📄 AI Estimator — generate`);
    console.log(`   Project: ${projectData.projectName}`);
    console.log(`   File: ${req.file.originalname} (${(req.file.size / 1024 / 1024).toFixed(2)} MB)`);

    // Call OpenAI to generate scope
    const rawResult = await generateScope(uploadedPath, projectData);

    // Sanitize scope to strip any prices/quantities
    const sanitized = {
      ...rawResult,
      divisions: sanitizeDivisions(rawResult.divisions || []),
    };

    console.log(`✓ Generated ${sanitized.divisions.length} division(s)\n`);

    return res.json(sanitized);

  } catch (error) {
    console.error('❌ Generate failed:', error.message);

    if (error instanceof z.ZodError) {
      return res.status(400).json({ error: 'Validation error', details: error.errors });
    }

    if (error instanceof multer.MulterError) {
      if (error.code === 'LIMIT_FILE_SIZE') {
        const limit = process.env.MAX_UPLOAD_MB || 50;
        return res.status(400).json({ error: `File too large (max ${limit} MB)` });
      }
      return res.status(400).json({ error: error.message });
    }

    return res.status(500).json({
      error: 'Failed to generate scope',
      message: error.message,
    });

  } finally {
    // Always delete uploaded PDF after processing
    if (uploadedPath) {
      try { await fs.unlink(uploadedPath); } catch {}
    }
  }
});

// ─────────────────────────────────────────────
// POST /api/ai-estimator/export
// ─────────────────────────────────────────────
router.post('/export', requireAuth, async (req, res) => {
  try {
    const { project, divisions } = exportBodySchema.parse(req.body);

    console.log(`\n📊 AI Estimator — export`);
    console.log(`   Project: ${project.name}`);

    const xlsxBuffer = await generateExcel(project, divisions);

    const filename = `ANDCON_SOW_${project.name.replace(/[^a-zA-Z0-9]/g, '_')}_v${project.version || '1'}.xlsx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Length', xlsxBuffer.byteLength);

    console.log(`✓ Export complete: ${filename}\n`);

    return res.send(Buffer.from(xlsxBuffer));

  } catch (error) {
    console.error('❌ Export failed:', error.message);

    if (error instanceof z.ZodError) {
      return res.status(400).json({ error: 'Validation error', details: error.errors });
    }

    return res.status(500).json({
      error: 'Failed to generate Excel',
      message: error.message,
    });
  }
});

export default router;
