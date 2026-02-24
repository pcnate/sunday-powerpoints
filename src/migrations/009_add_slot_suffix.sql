-- Migration: 009_add_slot_suffix
-- Description: Add optional letter suffix to song_selections for sub-slot identification (e.g., Song 2a, Song 2b).

ALTER TABLE song_selections
  ADD COLUMN slot_suffix CHAR(1) DEFAULT NULL AFTER slot_number;

DROP INDEX idx_date_slot ON song_selections;
CREATE INDEX idx_date_slot ON song_selections ( sunday_date, slot_type, slot_number, slot_suffix );
