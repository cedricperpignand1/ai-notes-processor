# AI Estimator — Setup & Deploy Guide

Generates a scope of work from construction plan PDFs in ANDCON proposal format.
PDF upload goes **directly to the processor** (bypasses Vercel entirely).

---

## Architecture

```
Browser → GET /api/ai-estimator/token (Vercel)   → short-lived HMAC token
Browser → POST {PROCESSOR_URL}/api/ai-estimator/generate  (Processor, multipart PDF)
Browser → POST {PROCESSOR_URL}/api/ai-estimator/export    (Processor, JSON → .xlsx)
```

No PDF or large file ever passes through Vercel serverless functions.

---

## Environment Variables

### Vercel (Next.js)
| Variable | Value | Notes |
|---|---|---|
| `NEXT_PUBLIC_PROCESSOR_BASE_URL` | `https://your-service.onrender.com` | Public — safe for browser |
| `PROCESSOR_SHARED_SECRET` | long random string | Server-only — signs HMAC tokens |

### Processor (Render / any Node host)
| Variable | Value | Notes |
|---|---|---|
| `OPENAI_API_KEY` | `sk-proj-...` | Required |
| `PROCESSOR_SHARED_SECRET` | same as Vercel | Must match exactly |
| `ALLOWED_ORIGINS` | `https://andconcm-modern.vercel.app,http://localhost:3000` | Comma-separated |
| `MAX_UPLOAD_MB` | `50` | Optional, default 50 |
| `PORT` | `3001` | Optional, default 3001 |

---

## Local Development

### 1. Start the processor
```bash
cd ai-notes-processor
npm install
cp .env.example .env
# Fill in OPENAI_API_KEY, PROCESSOR_SHARED_SECRET, ALLOWED_ORIGINS
node server.js
```

### 2. Start Next.js
```bash
cd andconcm-modern
npm install
# .env.local already has NEXT_PUBLIC_PROCESSOR_BASE_URL=http://localhost:3001
npm run dev
```

### 3. Open the page
```
http://localhost:3000/portal/ai-estimator
```

---

## Excel Template Setup

The processor loads your SOV template from:
```
ai-notes-processor/templates/SOV_template.xlsx
```

**Steps:**
1. Create the folder: `ai-notes-processor/templates/`
2. Copy your Excel SOV template into it, renamed to `SOV_template.xlsx`
3. The processor will write ONLY to:
   - `COVER!G26` → Project Name
   - `COVER!G27` → Project Address
   - `COVER!G28` → Version
   - `COVER!G29` → Date
   - `Detailed SOV` col B → Line Item Title (rows per division)
   - `Detailed SOV` col E → Scope (semicolon-separated)
4. All formulas in other cells are preserved

**If no template exists**, the processor generates a minimal fallback workbook
(no formulas). Use the real template for production.

---

## Deploy to Render

1. Push your `ai-notes-processor` folder to a Git repo (or Render can pull from a monorepo path)
2. Create a new **Web Service** on Render:
   - **Environment**: Node
   - **Build command**: `npm install`
   - **Start command**: `node server.js`
   - **Root directory**: `ai-notes-processor` (if monorepo)
3. Add all environment variables in the Render dashboard
4. Add your `templates/SOV_template.xlsx` via a persistent disk or bake it into the repo (recommended for MVP)
5. After deploy, copy your Render service URL (e.g. `https://ai-notes-processor.onrender.com`)

### Add URL to Vercel
In Vercel Dashboard → Project Settings → Environment Variables:
- `NEXT_PUBLIC_PROCESSOR_BASE_URL` = `https://ai-notes-processor.onrender.com`
- `PROCESSOR_SHARED_SECRET` = same value as Render

---

## How the AI Works

1. User uploads a PDF (plans, specifications, or both)
2. Processor reads the PDF directly using the OpenAI Responses API (`gpt-4o`)
3. OpenAI returns structured JSON with:
   - Divisions (DIV 2 through DIV 28, excluding DIV 1)
   - Line items with title + scope points
4. Server sanitizer strips any pricing/quantities that slipped through
5. User reviews and edits in the browser
6. Export writes to Excel template and downloads

### Scope rules enforced in prompt + sanitizer:
- No prices, dollar amounts, or costs
- No quantities (SF, LF, CY, etc.)
- No unit rates (per SF, per LF)
- No totals, lump sums, or allowances
- DIV 1 General Requirements is always skipped

---

## Division Capacity (template row limits)

| Division | Rows |
|---|---|
| DIV 2 Sitework | 2 |
| DIV 3 Concrete | 4 |
| DIV 4 Masonry | 2 |
| DIV 5 Metals | 6 |
| DIV 6 Woods & Plastics | 9 |
| DIV 7 Thermal & Moisture | 9 |
| DIV 8 Openings | 1 |
| DIV 9 Finishes | 4 |
| DIV 10 Specialties | 1 |
| DIV 11 Equipment | 2 |
| DIV 12 Furnishings | 2 |
| DIV 13 Special Construction | 2 |
| DIV 14 Conveying | 2 |
| DIV 21 Fire Suppression | 2 |
| DIV 22 Plumbing | 2 |
| DIV 23 HVAC | 1 |
| DIV 26 Electrical | 3 |
| DIV 27 Telecom | 2 |
| DIV 28 Safety/Security | 1 |

If the AI returns more items than capacity, merge them manually in the editor
or use the **↻ Clean** button on any scope cell.

---

## Troubleshooting

### CORS error in browser console
- Check `ALLOWED_ORIGINS` on the processor includes your Vercel domain exactly
- No trailing slash: `https://andconcm-modern.vercel.app` ✓ not `https://andconcm-modern.vercel.app/` ✗

### "NEXT_PUBLIC_PROCESSOR_BASE_URL not configured"
- Add the env var to Vercel dashboard and redeploy
- For local: it's already in `.env.local`

### "Only PDF files are accepted"
- File must be `application/pdf` MIME type
- Rename `.PDF` → `.pdf` if on Windows

### "File too large"
- Default max is 50 MB. Increase `MAX_UPLOAD_MB` on the processor if needed.

### Render free tier cold starts (30s+ delay)
- The processor may take 30+ seconds to respond on first request after inactivity
- The UI shows "AI is reading the plans…" — just wait
- Upgrade to Render Starter ($7/mo) to avoid cold starts

### Token expired
- Tokens are valid for 2 minutes. If generation takes > 2 minutes the token may
  expire, but this is only the auth token — the upload is already in progress.
  A new token is fetched fresh for each Generate and Export action.

### Excel formulas broken
- Never modify the Detailed SOV worksheet structure in the template
- Only copy data into existing rows; do not insert/delete rows
- The processor uses `exceljs` which preserves formulas in untouched cells

### OpenAI returns no divisions
- The PDF may be scanned images with no text layer — try a clearer or text-based PDF
- Check the processor logs for the OpenAI response
- Try a different `gpt-4o` prompt by editing `utils/openai-estimator.js`
