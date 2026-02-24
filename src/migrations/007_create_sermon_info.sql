-- Migration: 007_create_sermon_info
-- Description: Create table for storing sermon presentation info per Sunday folder

CREATE TABLE sermon_info (
  sunday_date  CHAR(8) NOT NULL PRIMARY KEY,
  pptx_file    VARCHAR(512) DEFAULT NULL,
  title        VARCHAR(512) DEFAULT NULL,
  speaker      VARCHAR(512) DEFAULT NULL,
  created_at   DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at   DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
