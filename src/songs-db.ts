import { getPool, healthCheck } from './db';
import { RowDataPacket } from 'mysql2/promise';


/**
 * Normalize a song name for fuzzy matching.
 * Strips extension, "- Shortcut" suffix, leading number prefix,
 * lowercases, replaces hyphens with spaces, removes non-alphanumeric
 * characters (except spaces/digits), collapses whitespace.
 *
 * @param name - raw song display name
 * @returns normalized string for comparison
 */
export function normalizeSongName( name: string ): string {
  let n = name;

  // Remove file extension
  n = n.replace( /\.(pptx?|ppt)$/i, '' );

  // Remove " - Shortcut" suffix
  n = n.replace( /\s*-\s*Shortcut$/i, '' );

  // Strip leading number prefix (e.g. "203 - " or "203 ")
  n = n.replace( /^\d+\s*[-–—]?\s*/, '' );

  // Lowercase
  n = n.toLowerCase();

  // Replace hyphens/dashes with spaces
  n = n.replace( /[-–—]/g, ' ' );

  // Remove non-alphanumeric (keep spaces and digits)
  n = n.replace( /[^a-z0-9\s]/g, '' );

  // Collapse whitespace and trim
  n = n.replace( /\s+/g, ' ' ).trim();

  return n;
}


/**
 * Find or create a song record using tiered matching.
 *
 * @param name - display name of the song
 * @param number - optional song number
 * @param book - optional book/folder name
 * @param filePath - optional relative file path within the songs directory
 * @returns the song ID
 */
export async function findOrCreateSong( name: string, number?: string, book?: string, filePath?: string ): Promise<number> {
  const pool = getPool();
  const normalized = normalizeSongName( name );

  // Clean display name: strip extension and number prefix (number stored separately)
  let displayName = name.replace( /\.(pptx?|ppt)$/i, '' );
  if ( number ) displayName = displayName.replace( /^\d+\s*[-–—]?\s*/, '' );
  displayName = displayName.trim();

  /**
   * Build and execute an UPDATE to backfill any missing fields on a matched record.
   *
   * @param id - the song record ID to update
   */
  const backfillFields = async ( id: number ) => {
    const updates: string[] = [];
    const params: any[] = [];
    if ( number ) { updates.push( 'number = COALESCE(number, ?)' ); params.push( number ); }
    if ( book ) { updates.push( 'book = COALESCE(book, ?)' ); params.push( book ); }
    if ( filePath ) { updates.push( 'file_path = COALESCE(file_path, ?)' ); params.push( filePath ); }
    if ( updates.length > 0 ) {
      params.push( id );
      await pool.query( `UPDATE songs SET ${ updates.join( ', ' ) } WHERE id = ?`, params );
    }
  };

  // Tier 1: number + normalized_name + book
  if ( number && book ) {
    const [ rows ] = await pool.query<RowDataPacket[]>(
      'SELECT id FROM songs WHERE number = ? AND normalized_name = ? AND book = ?',
      [ number, normalized, book ]
    );
    if ( rows.length === 1 ) {
      await backfillFields( rows[ 0 ].id );
      return rows[ 0 ].id;
    }
  }

  // Tier 2: number + normalized_name
  if ( number ) {
    const [ rows ] = await pool.query<RowDataPacket[]>(
      'SELECT id FROM songs WHERE number = ? AND normalized_name = ?',
      [ number, normalized ]
    );
    if ( rows.length === 1 ) {
      await backfillFields( rows[ 0 ].id );
      return rows[ 0 ].id;
    }
  }

  // Tier 3: normalized_name + book
  if ( book ) {
    const [ rows ] = await pool.query<RowDataPacket[]>(
      'SELECT id FROM songs WHERE normalized_name = ? AND book = ?',
      [ normalized, book ]
    );
    if ( rows.length === 1 ) {
      await backfillFields( rows[ 0 ].id );
      return rows[ 0 ].id;
    }
  }

  // Tier 4: normalized_name only
  const [ rows ] = await pool.query<RowDataPacket[]>(
    'SELECT id FROM songs WHERE normalized_name = ?',
    [ normalized ]
  );
  if ( rows.length === 1 ) {
    await backfillFields( rows[ 0 ].id );
    return rows[ 0 ].id;
  }

  // Tier 5: Insert new
  const [ result ] = await pool.query<any>(
    'INSERT INTO songs (name, normalized_name, number, book, file_path) VALUES (?, ?, ?, ?, ?)',
    [ displayName, normalized, number || null, book || null, filePath || null ]
  );
  return result.insertId;
}


/**
 * Record a song selection for a date/slot.
 * Soft-removes any existing active selection for the same slot first.
 *
 * @param songId - ID from the songs table
 * @param sundayDate - YYYYMMDD for songs/choruses, YYYYMM for closing songs
 * @param slotType - 'song', 'chorus', or 'closing'
 * @param slotNumber - slot number (1-3 for songs, 1-N for choruses, 1 for closing)
 * @param slotSuffix - optional letter suffix for split slots (e.g., 'a', 'b')
 */
export async function recordSongSelection(
  songId: number,
  sundayDate: string,
  slotType: 'song' | 'chorus' | 'closing',
  slotNumber: number,
  slotSuffix?: string
): Promise<void> {
  const pool = getPool();
  await removeSongSelection( sundayDate, slotType, slotNumber, slotSuffix );
  await pool.query(
    'INSERT INTO song_selections (song_id, sunday_date, slot_type, slot_number, slot_suffix) VALUES (?, ?, ?, ?, ?)',
    [ songId, sundayDate, slotType, slotNumber, slotSuffix || null ]
  );
}


/**
 * Soft-remove a song selection by setting removed_at.
 *
 * @param sundayDate - YYYYMMDD or YYYYMM
 * @param slotType - 'song', 'chorus', or 'closing'
 * @param slotNumber - slot number
 * @param slotSuffix - optional letter suffix for split slots (e.g., 'a', 'b')
 */
export async function removeSongSelection(
  sundayDate: string,
  slotType: 'song' | 'chorus' | 'closing',
  slotNumber: number,
  slotSuffix?: string
): Promise<void> {
  const pool = getPool();
  if ( slotSuffix ) {
    await pool.query(
      'UPDATE song_selections SET removed_at = NOW() WHERE sunday_date = ? AND slot_type = ? AND slot_number = ? AND slot_suffix = ? AND removed_at IS NULL',
      [ sundayDate, slotType, slotNumber, slotSuffix ]
    );
  } else {
    await pool.query(
      'UPDATE song_selections SET removed_at = NOW() WHERE sunday_date = ? AND slot_type = ? AND slot_number = ? AND slot_suffix IS NULL AND removed_at IS NULL',
      [ sundayDate, slotType, slotNumber ]
    );
  }
}


/**
 * Soft-remove all chorus selections for a date.
 *
 * @param sundayDate - YYYYMMDD
 */
export async function removeAllChorusSelections( sundayDate: string ): Promise<void> {
  const pool = getPool();
  await pool.query(
    'UPDATE song_selections SET removed_at = NOW() WHERE sunday_date = ? AND slot_type = ? AND removed_at IS NULL',
    [ sundayDate, 'chorus' ]
  );
}


/**
 * Get the active closing song for a month.
 *
 * @param year - the year
 * @param month - the month (1-12)
 * @returns song fields including ccli, or null if none set
 */
export async function getClosingSong( year: number, month: number ): Promise<{ id: number; name: string; number: string | null; book: string | null; ccli: string | null } | null> {
  const pool = getPool();
  const monthKey = `${ year }${ String( month ).padStart( 2, '0' ) }`;

  const [ rows ] = await pool.query<RowDataPacket[]>(
    `SELECT s.id, s.name, s.number, s.book, s.ccli
     FROM song_selections sel
     JOIN songs s ON s.id = sel.song_id
     WHERE sel.sunday_date = ? AND sel.slot_type = 'closing' AND sel.removed_at IS NULL
     ORDER BY sel.created_at DESC
     LIMIT 1`,
    [ monthKey ]
  );

  if ( rows.length === 0 ) return null;
  return { id: rows[ 0 ].id, name: rows[ 0 ].name, number: rows[ 0 ].number || null, book: rows[ 0 ].book || null, ccli: rows[ 0 ].ccli || null };
}


/**
 * Set or clear the closing song for a month.
 * When setting, also creates shortcuts in all Sunday folders for the month.
 *
 * @param year - the year
 * @param month - the month (1-12)
 * @param song - song to set, or null to clear
 * @returns the song ID if set, or null if cleared
 */
export async function setClosingSong(
  year: number,
  month: number,
  song: { name: string; number?: string; book?: string } | null
): Promise<number | null> {
  const monthKey = `${ year }${ String( month ).padStart( 2, '0' ) }`;

  // Always remove existing closing song for this month
  await removeSongSelection( monthKey, 'closing', 1 );

  if ( !song ) return null;

  // Find or create the song record, then record the selection
  const songId = await findOrCreateSong( song.name, song.number, song.book );
  await recordSongSelection( songId, monthKey, 'closing', 1 );
  return songId;
}


/**
 * Check whether an active selection already exists for a given date/slot.
 *
 * @param sundayDate - YYYYMMDD or YYYYMM
 * @param slotType - 'song', 'chorus', or 'closing'
 * @param slotNumber - slot number
 * @param slotSuffix - optional letter suffix for split slots (e.g., 'a', 'b')
 * @returns true if an active (non-removed) selection exists
 */
async function hasActiveSelection(
  sundayDate: string,
  slotType: 'song' | 'chorus' | 'closing',
  slotNumber: number,
  slotSuffix?: string
): Promise<boolean> {
  const pool = getPool();
  if ( slotSuffix ) {
    const [ rows ] = await pool.query<RowDataPacket[]>(
      'SELECT 1 FROM song_selections WHERE sunday_date = ? AND slot_type = ? AND slot_number = ? AND slot_suffix = ? AND removed_at IS NULL LIMIT 1',
      [ sundayDate, slotType, slotNumber, slotSuffix ]
    );
    return rows.length > 0;
  } else {
    const [ rows ] = await pool.query<RowDataPacket[]>(
      'SELECT 1 FROM song_selections WHERE sunday_date = ? AND slot_type = ? AND slot_number = ? AND slot_suffix IS NULL AND removed_at IS NULL LIMIT 1',
      [ sundayDate, slotType, slotNumber ]
    );
    return rows.length > 0;
  }
}


/**
 * Parsed shortcut info extracted from a folder scan.
 */
export interface ShortcutInfo {
  slotType: 'song' | 'chorus' | 'closing';
  slotNumber: number;
  slotSuffix?: string;
  songName: string;
  songNumber?: string;
}


/**
 * Parse a shortcut filename into slot info.
 * Handles patterns like:
 *   "song 1 203 - Tell Me the Story - Shortcut.lnk"
 *   "chorus 2 Amazing Grace - Shortcut.lnk"
 *   "closing song 203 - Tell Me the Story - Shortcut.lnk"
 *
 * @param filename - the .lnk filename
 * @returns parsed info, or null if not a recognized song/chorus/closing shortcut
 */
export function parseShortcutFilename( filename: string ): ShortcutInfo | null {
  const lower = filename.toLowerCase();
  if ( !lower.endsWith( '.lnk' ) ) return null;

  // Strip " - Shortcut.lnk" suffix
  let base = filename.replace( /\s*-\s*Shortcut\.lnk$/i, '' );

  // Closing song: "closing song [number - ] name"
  const closingMatch = base.match( /^closing song\s+(?:(\d+)\s*-\s*)?(.+)$/i );
  if ( closingMatch ) {
    return {
      slotType: 'closing',
      slotNumber: 1,
      songName: closingMatch[ 2 ].trim(),
      songNumber: closingMatch[ 1 ] || undefined,
    };
  }

  // Song: "song N[a-z] [number - ] name" (optional letter suffix for split slots)
  const songMatch = base.match( /^song\s+(\d+)([a-z])?\s+(?:(\d+)\s*-\s*)?(.+)$/i );
  if ( songMatch ) {
    return {
      slotType: 'song',
      slotNumber: parseInt( songMatch[ 1 ] ),
      slotSuffix: songMatch[ 2 ]?.toLowerCase() || undefined,
      songName: songMatch[ 4 ].trim(),
      songNumber: songMatch[ 3 ] || undefined,
    };
  }

  // Chorus: "chorus N [number - ] name"
  const chorusMatch = base.match( /^chorus\s+(\d+)\s+(?:(\d+)\s*-\s*)?(.+)$/i );
  if ( chorusMatch ) {
    return {
      slotType: 'chorus',
      slotNumber: parseInt( chorusMatch[ 1 ] ),
      songName: chorusMatch[ 3 ].trim(),
      songNumber: chorusMatch[ 2 ] || undefined,
    };
  }

  return null;
}


/**
 * Structured song info from a scanned SundayFolder for DB sync.
 */
export interface FolderSongInfo {
  slotType: 'song' | 'chorus' | 'closing';
  slotNumber: number;
  slotSuffix?: string;
  name: string;
  number?: string;
  book?: string;
  target?: string;
}


/**
 * Extract a relative file path from an absolute target path and songs directory.
 *
 * @param target - absolute path to the song file
 * @param songsDirectory - absolute path to the songs root directory
 * @returns relative path, or undefined if target is not under songsDirectory
 */
function extractRelativePath( target: string, songsDirectory: string ): string | undefined {
  if ( !target || !songsDirectory ) return undefined;
  const normalizedTarget = target.replace( /\\/g, '/' ).toLowerCase();
  const normalizedBase = songsDirectory.replace( /\\/g, '/' ).toLowerCase();
  if ( normalizedTarget.startsWith( normalizedBase ) ) {
    return target.slice( songsDirectory.length ).replace( /^[/\\]+/, '' );
  }
  return undefined;
}


/**
 * Sync song selections from shortcut files in a Sunday folder.
 * Only inserts records that don't already exist (idempotent).
 * Accepts either raw .lnk filenames or structured song data from the cache.
 *
 * @param sundayDate - YYYYMMDD date string for the folder
 * @param songs - array of structured song info from the folder
 * @param songsDirectory - optional songs directory root for resolving relative paths
 */
export async function syncSelectionsFromFolder(
  sundayDate: string,
  songs: FolderSongInfo[],
  songsDirectory?: string
): Promise<void> {
  for ( const info of songs ) {
    // For closing songs, use YYYYMM as the date key
    const dateKey = info.slotType === 'closing'
      ? sundayDate.slice( 0, 6 )
      : sundayDate;

    // Skip if already tracked
    const exists = await hasActiveSelection( dateKey, info.slotType, info.slotNumber, info.slotSuffix );
    if ( exists ) continue;

    // Extract relative file path if we have a target and songs directory
    const filePath = ( info.target && songsDirectory )
      ? extractRelativePath( info.target, songsDirectory )
      : undefined;

    // Find or create the song, then record the selection
    const songId = await findOrCreateSong( info.name, info.number, info.book, filePath );
    await recordSongSelection( songId, dateKey, info.slotType, info.slotNumber, info.slotSuffix );
  }
}
