import { spawn } from 'child_process';
import path from 'path';

/**
 * Extract audio from video file using ffmpeg
 * @param {string} videoPath - Path to input video file
 * @param {string} audioPath - Path to output audio file
 * @returns {Promise<void>}
 */
export function extractAudio(videoPath, audioPath) {
  return new Promise((resolve, reject) => {
    console.log(`🎵 Extracting audio: ${path.basename(videoPath)} → ${path.basename(audioPath)}`);

    const ffmpeg = spawn('ffmpeg', [
      '-i', videoPath,           // Input file
      '-vn',                      // No video
      '-acodec', 'libmp3lame',   // MP3 codec
      '-ar', '16000',             // 16kHz sample rate (optimal for Whisper)
      '-ac', '1',                 // Mono audio
      '-b:a', '64k',              // 64kbps bitrate
      '-y',                       // Overwrite output
      audioPath
    ]);

    let stderr = '';

    ffmpeg.stderr.on('data', (data) => {
      stderr += data.toString();
    });

    ffmpeg.on('close', (code) => {
      if (code === 0) {
        console.log(`✓ Audio extracted successfully`);
        resolve();
      } else {
        console.error('FFmpeg stderr:', stderr);
        reject(new Error(`FFmpeg exited with code ${code}`));
      }
    });

    ffmpeg.on('error', (err) => {
      reject(new Error(`FFmpeg spawn error: ${err.message}`));
    });
  });
}

/**
 * Check if ffmpeg is installed
 */
export function checkFFmpeg() {
  return new Promise((resolve) => {
    const check = spawn('ffmpeg', ['-version']);

    check.on('close', (code) => {
      resolve(code === 0);
    });

    check.on('error', () => {
      resolve(false);
    });
  });
}
