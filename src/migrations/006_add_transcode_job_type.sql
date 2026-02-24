-- Migration: 006_add_transcode_job_type
-- Description: Add 'transcode' to the jobs type ENUM for Kdenlive rendering pipeline

ALTER TABLE jobs MODIFY COLUMN type
  ENUM('video-alignment', 'transcription', 'claude-processing', 'transcode') NOT NULL;
