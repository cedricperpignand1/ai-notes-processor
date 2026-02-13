#!/usr/bin/env bash
# Render build script — installs FFmpeg + Node dependencies
set -e

apt-get update && apt-get install -y --no-install-recommends ffmpeg
npm ci --omit=dev
