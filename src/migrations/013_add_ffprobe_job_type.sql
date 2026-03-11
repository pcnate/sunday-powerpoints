ALTER TABLE jobs MODIFY COLUMN type ENUM('transcription','claude-processing','transcode','ffprobe') NOT NULL;
