import OpenAI from 'openai';
import fs from 'fs';
import path from 'path';

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
              'DIV 14', 'DIV 21', 'DIV 22', 'DIV 23', 'DIV 26', 'DIV 27', 'DIV 28',
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
7. Scope points must be concise, factual, plan-based observations.
8. If a detail is unclear or not shown, write "Verify in plans".
9. Scope points must be short phrases (not full sentences).

## DIVISION CAPACITY LIMITS (merge similar items if over capacity):
${capacityLines}

## PROJECT INFO:
- Project Name: ${projectName}
- Project Address: ${projectAddress}

Return the JSON schema exactly as specified. No extra keys.`;
}

/**
 * Generate scope of work from a PDF file using OpenAI.
 * @param {string} pdfPath - absolute path to the uploaded PDF
 * @param {{ projectName, projectAddress, version, date }} projectData
 * @returns {Promise<object>} structured scope JSON
 */
export async function generateScope(pdfPath, projectData) {
  const { projectName, projectAddress, version, date } = projectData;

  // Read PDF as base64
  const pdfBuffer = fs.readFileSync(pdfPath);
  const base64Pdf = pdfBuffer.toString('base64');

  const systemPrompt = buildSystemPrompt(projectName, projectAddress);

  // Use the OpenAI Responses API which supports direct PDF input
  // Falls back to chat completions with extracted text if responses API unavailable
  let result;

  try {
    // Try Responses API (OpenAI SDK ≥ 4.86)
    const response = await openai.responses.create({
      model: 'gpt-4o',
      input: [
        {
          role: 'user',
          content: [
            {
              type: 'input_file',
              filename: path.basename(pdfPath),
              file_data: `data:application/pdf;base64,${base64Pdf}`,
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
    // Fallback: Chat Completions API (works with all SDK versions)
    console.log('Responses API unavailable, falling back to chat completions:', responsesErr.message);

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
            // Try file content type (SDK ≥ 4.72)
            {
              type: 'file',
              file: {
                filename: path.basename(pdfPath),
                data: base64Pdf,
                mime_type: 'application/pdf',
              },
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
      max_tokens: 4096,
    });

    result = JSON.parse(completion.choices[0].message.content);
  }

  // Overwrite project fields with the user-provided values (not AI guesses)
  result.project = {
    name:    projectName    || result.project?.name    || '',
    address: projectAddress || result.project?.address || '',
    version: version        || result.project?.version || '1.0',
    date:    date           || result.project?.date    || new Date().toISOString().split('T')[0],
  };

  return result;
}
