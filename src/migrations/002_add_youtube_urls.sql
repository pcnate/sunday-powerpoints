-- Migration: 002_add_youtube_urls
-- Description: Create table for storing YouTube video URLs per Sunday folder

CREATE TABLE youtube_urls (
  sunday_date CHAR(8) NOT NULL PRIMARY KEY,
  url         VARCHAR(512) NOT NULL,
  created_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
