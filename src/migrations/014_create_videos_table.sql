CREATE TABLE IF NOT EXISTS videos (
  id                INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
  input_path        VARCHAR(768) NOT NULL,
  sunday_date       CHAR(8) NOT NULL,
  filename          VARCHAR(255) NOT NULL,
  status            ENUM('pending','completed','failed') NOT NULL DEFAULT 'pending',
  job_id            INT DEFAULT NULL,
  duration_secs     DOUBLE DEFAULT NULL,
  duration_timecode VARCHAR(20) DEFAULT NULL,
  framerate         DOUBLE DEFAULT NULL,
  width             INT UNSIGNED DEFAULT NULL,
  height            INT UNSIGNED DEFAULT NULL,
  video_codec       VARCHAR(50) DEFAULT NULL,
  audio_codec       VARCHAR(50) DEFAULT NULL,
  audio_streams     TINYINT UNSIGNED DEFAULT NULL,
  bitrate           BIGINT UNSIGNED DEFAULT NULL,
  file_size         BIGINT UNSIGNED DEFAULT NULL,
  file_created_at   DATETIME DEFAULT NULL,
  created_at        DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at        DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  UNIQUE INDEX idx_videos_input_path (input_path),
  INDEX idx_videos_sunday_date (sunday_date),
  INDEX idx_videos_status (status)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

-- Backfill from existing completed ffprobe jobs
INSERT INTO videos (input_path, sunday_date, filename, status, job_id,
  duration_secs, duration_timecode, framerate, width, height,
  video_codec, audio_codec, audio_streams, bitrate, file_size, file_created_at)
SELECT
  j.input_path,
  j.sunday_date,
  COALESCE(JSON_UNQUOTE(JSON_EXTRACT(j.metadata, '$.filename')), SUBSTRING_INDEX(j.input_path, '/', -1)),
  'completed',
  j.id,
  JSON_EXTRACT(j.metadata, '$.duration_secs'),
  JSON_UNQUOTE(JSON_EXTRACT(j.metadata, '$.duration_timecode')),
  JSON_EXTRACT(j.metadata, '$.framerate'),
  JSON_EXTRACT(j.metadata, '$.width'),
  JSON_EXTRACT(j.metadata, '$.height'),
  JSON_UNQUOTE(JSON_EXTRACT(j.metadata, '$.video_codec')),
  JSON_UNQUOTE(JSON_EXTRACT(j.metadata, '$.audio_codec')),
  JSON_EXTRACT(j.metadata, '$.audio_streams'),
  JSON_EXTRACT(j.metadata, '$.bitrate'),
  JSON_EXTRACT(j.metadata, '$.file_size'),
  JSON_UNQUOTE(JSON_EXTRACT(j.metadata, '$.created_at'))
FROM jobs j
WHERE j.type = 'ffprobe' AND j.status = 'completed' AND j.metadata IS NOT NULL
ON DUPLICATE KEY UPDATE videos.id = videos.id;
