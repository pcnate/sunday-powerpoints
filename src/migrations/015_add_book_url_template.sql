-- Migration: 015_add_book_url_template
-- Description: Add url_template column to books table for linking to online hymnals.
-- The template uses {number} as a placeholder for the song number (leading zeros stripped).
-- Example: https://hymnary.org/hymn/POSH1979/{number}

ALTER TABLE books ADD COLUMN url_template VARCHAR(512) DEFAULT NULL AFTER icon;
