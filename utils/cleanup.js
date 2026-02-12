import fs from 'fs/promises';
import path from 'path';

/**
 * Delete a file if it exists, suppress errors
 */
export async function deleteFile(filePath) {
  try {
    await fs.unlink(filePath);
    console.log(`✓ Deleted: ${path.basename(filePath)}`);
  } catch (err) {
    if (err.code !== 'ENOENT') {
      console.warn(`⚠ Failed to delete ${path.basename(filePath)}:`, err.message);
    }
  }
}

/**
 * Cleanup temp files (video + audio) for a given jobId
 */
export async function cleanupTempFiles(jobId) {
  const tmpDir = path.join(process.cwd(), 'tmp');
  const videoPath = path.join(tmpDir, `${jobId}.mp4`);
  const audioPath = path.join(tmpDir, `${jobId}.mp3`);

  await Promise.all([
    deleteFile(videoPath),
    deleteFile(audioPath)
  ]);
}

/**
 * Generate a unique job ID
 */
export function generateJobId() {
  return `job_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}
