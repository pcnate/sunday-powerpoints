-- Migration: 001_create_jobs_tables
-- Description: Create job queue and job logs tables

CREATE TABLE IF NOT EXISTS migrations (
  id         INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
  name       VARCHAR(255) NOT NULL UNIQUE,
  applied_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE jobs (
  id            INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
  type          ENUM('video-alignment', 'transcription', 'claude-processing') NOT NULL,
  status        ENUM('pending', 'queued', 'processing', 'completed', 'failed', 'cancelled') NOT NULL DEFAULT 'pending',
  priority      TINYINT UNSIGNED NOT NULL DEFAULT 5,
  sunday_date   CHAR(8) NOT NULL,
  input_path    VARCHAR(1024) NOT NULL,
  output_path   VARCHAR(1024) DEFAULT NULL,
  metadata      JSON DEFAULT NULL,
  error_message TEXT DEFAULT NULL,
  worker_id     VARCHAR(64) DEFAULT NULL,
  retry_count   TINYINT UNSIGNED NOT NULL DEFAULT 0,
  max_retries   TINYINT UNSIGNED NOT NULL DEFAULT 3,
  parent_job_id INT UNSIGNED DEFAULT NULL,
  created_at    DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at    DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  started_at    DATETIME DEFAULT NULL,
  completed_at  DATETIME DEFAULT NULL,
  heartbeat_at  DATETIME DEFAULT NULL,
  INDEX idx_status_type_priority (status, type, priority),
  INDEX idx_sunday_date (sunday_date),
  INDEX idx_parent_job_id (parent_job_id),
  INDEX idx_status_heartbeat (status, heartbeat_at),
  CONSTRAINT fk_parent_job FOREIGN KEY (parent_job_id) REFERENCES jobs(id) ON DELETE SET NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE job_logs (
  id         BIGINT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
  job_id     INT UNSIGNED NOT NULL,
  level      ENUM('info', 'warn', 'error', 'debug') NOT NULL DEFAULT 'info',
  message    TEXT NOT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  INDEX idx_job_id (job_id),
  CONSTRAINT fk_job_logs_job FOREIGN KEY (job_id) REFERENCES jobs(id) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
