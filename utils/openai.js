import OpenAI from 'openai';
import fs from 'fs';

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

/**
 * Transcribe audio file using OpenAI Whisper
 * @param {string} audioPath - Path to audio file
 * @returns {Promise<string>} - Transcribed text
 */
export async function transcribeAudio(audioPath) {
  console.log(`🎤 Transcribing audio...`);

  try {
    const audioFile = fs.createReadStream(audioPath);

    const response = await openai.audio.transcriptions.create({
      file: audioFile,
      model: 'whisper-1',
      language: 'en',
    });

    console.log(`✓ Transcription complete (${response.text.length} chars)`);
    return response.text;
  } catch (error) {
    console.error('Transcription error:', error.message);
    throw new Error(`Transcription failed: ${error.message}`);
  }
}

/**
 * Generate structured daily notes from transcript
 * @param {string} transcript - The transcribed text
 * @param {string} projectName - Project name
 * @param {string} createdBy - User who created the log
 * @returns {Promise<string>} - Structured notes
 */
export async function generateNotes(transcript, projectName, createdBy) {
  console.log(`📝 Generating structured notes...`);

  const systemPrompt = "You write concise, professional construction daily logs for ANDCON.";

  const userPrompt = `Project: ${projectName}
Created By: ${createdBy}
Transcript:
${transcript}

Write a Daily Log in EXACTLY this format with bullet points and short lines:

Work Completed Today:
- ...

In Progress:
- ...

Plan Next 24–48 Hours:
- ...

Issues / Blockers / RFIs Needed:
- ...

Deliveries / Inspections:
- ...

Safety / Access Notes:
- ...

Action Items:
- Owner: ...
- PM: ...
- Sub: ...
- Vendor: ...

Rules:
- If a section is not mentioned, write 'Not mentioned' under that section.
- Keep language clear and jobsite-real.
- Do not add assumptions.`;

  try {
    const response = await openai.chat.completions.create({
      model: 'gpt-4o-mini',
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      temperature: 0.3,
      max_tokens: 2000,
    });

    const notes = response.choices[0].message.content.trim();
    console.log(`✓ Notes generated (${notes.length} chars)`);
    return notes;
  } catch (error) {
    console.error('Notes generation error:', error.message);
    throw new Error(`Notes generation failed: ${error.message}`);
  }
}
