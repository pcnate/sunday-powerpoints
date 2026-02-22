-- Migration: 002_create_bible_verses_table
-- Description: Create table for storing bible verse references per month

CREATE TABLE bible_verses (
  id         INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
  year       SMALLINT UNSIGNED NOT NULL,
  month      TINYINT UNSIGNED NOT NULL,
  reference  VARCHAR(255) NOT NULL,
  created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  UNIQUE KEY idx_year_month (year, month)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
