-- Migration: 010_add_book_icon
-- Description: Add icon column to books table for custom book icon images.
-- The icon column stores the filename (e.g. 'awana.png') of an image file
-- located in SONGS_DIRECTORY/icons/. NULL means no icon.

ALTER TABLE books ADD COLUMN icon VARCHAR(255) DEFAULT NULL AFTER name;
