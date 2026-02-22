-- Migration: 004_create_songs_tables
-- Description: Create song registry and selection history tables for tracking
-- which songs are used each Sunday. Supports songs, choruses, and closing songs.

CREATE TABLE songs (
  id              INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
  name            VARCHAR(255) NOT NULL,
  normalized_name VARCHAR(255) NOT NULL,
  number          VARCHAR(10) DEFAULT NULL,
  book            VARCHAR(255) DEFAULT NULL,
  file_path       VARCHAR(1024) DEFAULT NULL,
  ccli            VARCHAR(20) DEFAULT NULL,
  created_at      DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at      DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  INDEX idx_normalized_name (normalized_name),
  INDEX idx_number_name (number, normalized_name),
  INDEX idx_ccli (ccli)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE song_selections (
  id          INT UNSIGNED AUTO_INCREMENT PRIMARY KEY,
  song_id     INT UNSIGNED NOT NULL,
  sunday_date CHAR(8) NOT NULL,
  slot_type   ENUM('song','chorus','closing') NOT NULL,
  slot_number TINYINT UNSIGNED NOT NULL DEFAULT 1,
  created_at  DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  removed_at  DATETIME DEFAULT NULL,
  INDEX idx_sunday_date (sunday_date),
  INDEX idx_song_id (song_id),
  INDEX idx_date_slot (sunday_date, slot_type, slot_number),
  CONSTRAINT fk_song_selections_song FOREIGN KEY (song_id) REFERENCES songs(id)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
