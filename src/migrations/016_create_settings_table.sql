CREATE TABLE IF NOT EXISTS settings (
  `key` VARCHAR(100) PRIMARY KEY,
  `value` TEXT NOT NULL,
  updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
);

INSERT INTO settings (`key`, `value`) VALUES ('claude-prompt', 'Read the VTT file and provide:

1. Sermon title
2. YouTube timestamp bookmarks for: welcome, songs (with hymn number), memory verse (with reference), tithes & offering, scripture reading, opening prayer, sermon start, and closing song
3. List all Bible verses referenced during the sermon using {book} {chapter} v{verse} format
4. Suggested YouTube tags for the sermon');
