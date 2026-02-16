import OpenAI, { toFile } from 'openai';
import fs from 'fs';

const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });

// ─────────────────────────────────────────────
// Structured output schema
// ─────────────────────────────────────────────
const SCOPE_SCHEMA = {
  type: 'object',
  properties: {
    project: {
      type: 'object',
      properties: {
        name:    { type: 'string' },
        address: { type: 'string' },
        version: { type: 'string' },
        date:    { type: 'string' },
      },
      required: ['name', 'address', 'version', 'date'],
      additionalProperties: false,
    },
    divisions: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          division: {
            type: 'string',
            enum: [
              'DIV 2', 'DIV 3', 'DIV 4', 'DIV 5', 'DIV 6', 'DIV 7',
              'DIV 8', 'DIV 9', 'DIV 10', 'DIV 11', 'DIV 12', 'DIV 13',
              'DIV 14', 'DIV 21', 'DIV 22', 'DIV 23', 'DIV 26', 'DIV 27', 'DIV 28', 'DIV 32',
            ],
          },
          items: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                title:        { type: 'string' },
                scope_points: { type: 'array', items: { type: 'string' } },
              },
              required: ['title', 'scope_points'],
              additionalProperties: false,
            },
          },
        },
        required: ['division', 'items'],
        additionalProperties: false,
      },
    },
    notes: { type: 'string' },
  },
  required: ['project', 'divisions', 'notes'],
  additionalProperties: false,
};

// Division capacities — used in prompt so AI knows to merge when needed
const DIVISION_CAPACITIES = {
  'DIV 2':  2,  'DIV 3':  4,  'DIV 4':  2,  'DIV 5':  6,
  'DIV 6':  9,  'DIV 7':  9,  'DIV 8':  1,  'DIV 9':  4,
  'DIV 10': 1,  'DIV 11': 2,  'DIV 12': 2,  'DIV 13': 2,
  'DIV 14': 2,  'DIV 21': 2,  'DIV 22': 2,  'DIV 23': 1,
  'DIV 26': 3,  'DIV 27': 2,  'DIV 28': 1,
  'DIV 32': 3,
};

function buildSystemPrompt(projectName, projectAddress) {
  const capacityLines = Object.entries(DIVISION_CAPACITIES)
    .map(([div, cap]) => `  ${div}: max ${cap} line item${cap > 1 ? 's' : ''}`)
    .join('\n');

  return `You are a professional construction estimator for ANDCON Construction Management.

Your task: Review the attached construction documents and produce a structured scope of work in CSI MasterFormat.

## HARD RULES — YOU MUST FOLLOW THESE EXACTLY:
1. DO NOT include any pricing, dollar amounts, costs, or budget figures.
2. DO NOT include any quantities (SF, LF, CY, tons, ea, etc.).
3. DO NOT include unit rates (per SF, per LF, ea @, etc.).
4. DO NOT include totals, subtotals, allowances, or lump sums.
5. SKIP Division 1 (General Requirements) entirely — it is done manually.
6. Only list divisions where work is clearly shown or specified in the plans.
7. Each scope point must be written as a COMPREHENSIVE NARRATIVE PARAGRAPH that describes the FULL SCOPE OF WORK for that item — not just a short bullet or spec callout. Each scope point should read like a complete subcontractor scope description that covers:
   - What is being installed, furnished, or constructed
   - How it connects to adjacent systems or structures
   - Key technical specifications found in the plans: model numbers, manufacturer names, pipe sizes, material specs, equipment ratings, dimensions, gauges, thicknesses, grades, finish types, colors, connection types, fasteners, and methods
   - Required testing, inspections, flushing, or commissioning
   - Coordination with other trades (penetrations, chases, sleeves, fire-stopping, backing, blocking)
   - Any specialty items, accessories, or ancillary work shown on the plans

   EXAMPLE of the level of detail expected for a single scope point (Plumbing):
   "Install new hot and cold water piping throughout the residence per plans; connect to municipal water service with main shutoff valve and backflow preventer; distribute water to all fixtures including toilets, sinks, showers, tubs, kitchen, and laundry; install water heater(s) complete with safety pan, T&P relief valve, and discharge piping; install sanitary drainage piping underground and above ground, connect to sewer, and provide proper slope, cleanouts, and vent stacks through roof; furnish and install all specified plumbing fixtures with necessary valves, drains, and trim; provide specialty items such as floor drains with trap primers, sump pump, or ejector system if shown; perform pressure tests, flush and sanitize the water system; coordinate with other trades for penetrations and chases, and ensure all required sleeves and fire-stopping are installed."

8. If a detail is unclear or not shown, write "Verify in plans".
9. Provide as many scope points as needed to fully describe the work for each line item. Each scope point should be a rich, flowing paragraph — not a short phrase. More detail is always better. Cover every aspect of the work: materials, installation methods, connections, testing, and trade coordination.

## DIVISION CAPACITY LIMITS (merge similar items if over capacity):
${capacityLines}

## PROJECT INFO:
- Project Name: ${projectName}
- Project Address: ${projectAddress}

Return the JSON schema exactly as specified. No extra keys.`;
}

/**
 * Upload PDF to OpenAI via Files API using a read stream (never loads full file into memory).
 * Uses toFile() to attach the correct filename + MIME type so OpenAI recognises it as a PDF.
 * @param {string} pdfPath - absolute path to the PDF on disk
 * @returns {Promise<string>} the OpenAI file ID
 */
async function uploadPdfToOpenAI(pdfPath) {
  // Extract the filename that multer saved (e.g. "1707123456-Plans.pdf")
  const filename = pdfPath.split(/[/\\]/).pop();

  const file = await openai.files.create({
    file: await toFile(fs.createReadStream(pdfPath), filename, { type: 'application/pdf' }),
    purpose: 'user_data',
  });
  console.log(`  Uploaded to OpenAI: ${file.id} (${file.bytes} bytes)`);
  return file.id;
}

/**
 * Delete a file from OpenAI (best-effort cleanup).
 * @param {string} fileId
 */
async function deleteOpenAIFile(fileId) {
  if (!fileId) return;
  try {
    await openai.files.del(fileId);
    console.log(`  Deleted OpenAI file: ${fileId}`);
  } catch (err) {
    console.warn(`  Failed to delete OpenAI file ${fileId}:`, err.message);
  }
}

/**
 * Generate scope of work from a PDF file using OpenAI.
 * Uses the Files API to stream the PDF from disk — never loads the entire file into memory.
 *
 * @param {string} pdfPath - absolute path to the uploaded PDF
 * @param {{ projectName, projectAddress, version, date }} projectData
 * @returns {Promise<{ result: object, fileId: string|null }>} structured scope JSON + file ID for cleanup
 */
export async function generateScope(pdfPath, projectData) {
  const { projectName, projectAddress, version, date } = projectData;
  const systemPrompt = buildSystemPrompt(projectName, projectAddress);

  // Step 1: Stream-upload PDF to OpenAI (memory-safe)
  const fileId = await uploadPdfToOpenAI(pdfPath);

  let result;

  try {
    // Step 2: Use Responses API with file reference (no base64 in memory)
    const response = await openai.responses.create({
      model: 'gpt-4o',
      input: [
        {
          role: 'user',
          content: [
            {
              type: 'input_file',
              file_id: fileId,
            },
            {
              type: 'input_text',
              text: systemPrompt,
            },
          ],
        },
      ],
      text: {
        format: {
          type: 'json_schema',
          name: 'scope_of_work',
          schema: SCOPE_SCHEMA,
          strict: true,
        },
      },
    });

    result = JSON.parse(response.output_text);
  } catch (responsesErr) {
    console.log('  Responses API failed, trying Chat Completions fallback:', responsesErr.message);

    // Fallback: Chat Completions with file reference
    const completion = await openai.chat.completions.create({
      model: 'gpt-4o',
      messages: [
        { role: 'system', content: systemPrompt },
        {
          role: 'user',
          content: [
            {
              type: 'text',
              text: 'Analyze the attached construction plans PDF and generate a structured scope of work.',
            },
            {
              type: 'file',
              file: { file_id: fileId },
            },
          ],
        },
      ],
      response_format: {
        type: 'json_schema',
        json_schema: {
          name: 'scope_of_work',
          strict: true,
          schema: SCOPE_SCHEMA,
        },
      },
      max_tokens: 16384,
    });

    result = JSON.parse(completion.choices[0].message.content);
  } finally {
    // Step 3: Delete the uploaded file from OpenAI (best-effort)
    await deleteOpenAIFile(fileId);
  }

  // Overwrite project fields with user-provided values (not AI guesses)
  result.project = {
    name:    projectName    || result.project?.name    || '',
    address: projectAddress || result.project?.address || '',
    version: version        || result.project?.version || '1.0',
    date:    date           || result.project?.date    || new Date().toISOString().split('T')[0],
  };

  return result;
}
