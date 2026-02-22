# Plan: Database-Backed Song Tracking

## Context

Songs are currently tracked only via filesystem shortcuts (.lnk files) in Sunday folders. There's no persistent record of which songs were used when, making statistical analysis and CCLI reporting impossible. Songs can also be renamed (case, spacing) or moved between folders in OneDrive, which could break any future identity-based tracking.

This plan adds a MySQL song registry and selection history table alongside the existing shortcut system. Shortcuts remain the functional layer; the database is the historical/statistical layer.

## Database Schema

**New migration: `src/migrations/004_create_songs_tables.sql`**

### `songs` table (master registry)
| Column | Type | Purpose |
|--------|------|---------|
| `id` | INT UNSIGNED AUTO_INCREMENT PK | Stable unique identifier |
| `name` | VARCHAR(255) | Display name (e.g., "Tell Me the Story") |
| `normalized_name` | VARCHAR(255) | Lowercased, whitespace-collapsed, punctuation-stripped — for matching |
| `number` | VARCHAR(10) NULL | Number prefix from filename (e.g., "203") |
| `book` | VARCHAR(255) NULL | Source folder/collection |
| `file_path` | VARCHAR(1024) NULL | Current relative path within SONGS_DIRECTORY |
| `ccli` | VARCHAR(20) NULL | CCLI number (for future reporting, populated manually) |
| `created_at` / `updated_at` | DATETIME | Timestamps |

Indexes: `normalized_name`, `(number, normalized_name)`, `ccli`

### `song_selections` table (usage tracking)
| Column | Type | Purpose |
|--------|------|---------|
| `id` | INT UNSIGNED AUTO_INCREMENT PK | Row ID |
| `song_id` | INT UNSIGNED FK→songs | Which song |
| `sunday_date` | CHAR(8) | YYYYMMDD |
| `slot_type` | ENUM('song','chorus','closing') | Song slot, chorus slot, or closing song |
| `slot_number` | TINYINT UNSIGNED | 1-3 for songs, 1-N for choruses |
| `created_at` | DATETIME | When selected |
| `removed_at` | DATETIME NULL | When deselected (NULL = still active) |

Indexes: `sunday_date`, `song_id`, `(sunday_date, slot_type, slot_number)`

Songs, choruses, and closing songs share the `songs` table — a chorus or closing song is just a song used in a different slot type. Closing songs are per-month: their `sunday_date` is `YYYYMM` (6 chars) rather than `YYYYMMDD`. A closing song shortcut is created in every Sunday folder for that month.

## Song Identity & Matching

### Normalization function
Converts display names to a canonical form for fuzzy matching:
1. Remove file extension and " - Shortcut" suffix
2. Strip leading number prefix (stored separately)
3. Lowercase
4. Replace hyphens/dashes with spaces
5. Remove non-alphanumeric characters (except spaces and digits)
6. Collapse whitespace, trim

Example: `"203 - Tell Me the Story of Jesus.pptx"` → `"tell me the story of jesus"`

### Find-or-create algorithm (tiered matching)
When a song is selected, resolve it to a `songs.id`:

1. **number + normalized_name + book** → exact match (highest confidence)
2. **number + normalized_name** (ignoring book) → song moved between folders; update book
3. **normalized_name + book** → song has no number; match within same book
4. **normalized_name only** → 1 match = same song renamed/moved; update metadata
5. **No match** → INSERT new record

Each tier requires exactly 1 match to proceed. Multiple matches = skip to next tier for more specificity. This handles:
- Renames (case, spacing) via normalization
- Folder moves via path-independent matching
- Simultaneous rename+move gracefully falls through to insert (rare, can merge later)

## Files to Create

### `src/migrations/004_create_songs_tables.sql`
The two CREATE TABLE statements above.

### `src/songs-db.ts`
Utility module (following the pattern of `src/routes/jobs.ts`):
- `normalizeSongName( name )` — normalization function
- `findOrCreateSong( name, number?, book? )` — tiered matching algorithm
- `recordSongSelection( songId, sundayDate, slotType, slotNumber )` — insert selection (soft-removes previous selection for same slot first)
- `removeSongSelection( sundayDate, slotType, slotNumber )` — set `removed_at` on active selection
- `removeAllChorusSelections( sundayDate )` — bulk soft-remove choruses for a date
- `backfillFromShortcuts( outputDirectory, songsDirectory )` — scan existing folders for historical data

## Files to Modify

### `src/server.ts`

**`/api/update-song` endpoint (line 1511)**
After the shortcut is created successfully (line 1542), add:
- Call `findOrCreateSong()` with the song's name, number, and book (book extracted from the resolved `songFile` path)
- Call `removeSongSelection()` then `recordSongSelection()` for the slot
- If song was cleared (empty name), only call `removeSongSelection()`
- Entire DB block wrapped in try/catch — failures log but never break the shortcut workflow
- Guarded by `healthCheck()` — skipped if DB is unavailable

**`/api/chorus-update` endpoint (line 1558)**
After all chorus shortcuts are recreated (line 1603), add:
- Call `removeAllChorusSelections( date )` to soft-remove previous chorus records
- Loop through choruses, call `findOrCreateSong()` + `recordSongSelection()` for each
- Same try/catch + healthCheck guard

**New endpoint: `POST /api/admin/backfill-songs`**
- Calls `backfillFromShortcuts()` to scan existing Sunday folders
- Parses song/chorus shortcut filenames using the same regex patterns as `/api/week-song-selections`
- Idempotent — skips date+slot combinations that already have active records
- Returns count of folders scanned and selections recorded

### `src/types/song.ts` (new)
TypeScript interfaces: `SongRecord`, `SongSelection`

## Key Design Decisions

- **DB failure is non-fatal**: The shortcut system is the source of truth. All DB writes are guarded by `healthCheck()` and wrapped in try/catch. If MySQL is down, the app works exactly as before.
- **Single songs table for songs and choruses**: They're the same PowerPoint files from the same library; the distinction is only in `slot_type`.
- **Soft deletes via `removed_at`**: Preserves full selection history for statistics/CCLI reporting. "Which songs did we sing on 20260222?" = all records where `sunday_date = '20260222' AND removed_at IS NULL`.
- **No file content hashing**: Adds complexity for minimal benefit. The tiered name-matching handles 99% of real-world rename/move scenarios. If a rare edge case creates a duplicate, records can be merged manually later.
- **Book extracted from resolved path**: When `/api/update-song` calls `findSongFile()`, the returned path contains the book folder. We extract it from there rather than relying solely on the frontend's `book` field, giving us a more reliable value.
- **Closing songs are per-month**: Stored with `sunday_date = YYYYMM` and `slot_type = 'closing'`. A closing song shortcut is created in every Sunday folder for that month.

## Verification

1. **Start dev server**: `npm run dev` — migration 004 should auto-apply
2. **Verify tables created**: Connect to MySQL and confirm `songs` and `song_selections` tables exist
3. **Select a song in the UI**: Use the Planning tab to assign a song to a slot → verify a row appears in `songs` and `song_selections`
4. **Change the song**: Select a different song for the same slot → verify the old selection has `removed_at` set and a new selection row exists
5. **Clear a song slot**: Remove a song → verify `removed_at` is set, no new selection created
6. **Update choruses**: Add/remove choruses → verify chorus selections tracked with `slot_type = 'chorus'`
7. **Backfill**: Call `POST /api/admin/backfill-songs` → verify historical selections populated from existing shortcut files
8. **DB offline**: Stop MySQL, select a song in the UI → verify shortcuts still created, error logged but not surfaced to user
9. **Run tests**: `npm test` — existing tests should still pass
10. **New tests**: `npx jest tests/songs-db.test.ts` — normalization and matching logic tests
