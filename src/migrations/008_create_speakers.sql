-- Migration: 008_create_speakers
-- Description: Create table for storing reusable speaker names

CREATE TABLE speakers (
  id         INT AUTO_INCREMENT PRIMARY KEY,
  name       VARCHAR(256) NOT NULL UNIQUE,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
