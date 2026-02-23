/**
 * One-time script to scan Sunday folders and generate SQL INSERT
 * statements for the `songs` and `song_selections` tables.
 *
 * Reads .lnk shortcut files from each YYYYMMDD subfolder, resolves
 * their targets, and writes a .sql file organized by Sunday with
 * comments for each song.
 *
 * Usage:
 *   npx ts-node scripts/generate-song-inserts.ts <folder>
 *
 * Example:
 *   npx ts-node scripts/generate-song-inserts.ts "C:\Users\pcnate\OneDrive\PowerPoints\Sermon Videos\Sermons"
 *
 * Output:
 *   song-inserts.sql  (written into the scanned folder)
 */

import path from 'node:path';
import fs from 'node:fs';
import ws from 'windows-shortcuts';


// ── CLI argument ────────────────────────────────────────────────────

const scanDir = process.argv[ 2 ];

if ( !scanDir ) {
  console.error( 'Usage: npx ts-node scripts/generate-song-inserts.ts <folder>' );
  console.error( '  <folder>  Path to directory containing YYYYMMDD subfolders' );
  process.exit( 1 );
}

const resolvedDir = path.resolve( scanDir );

if ( !fs.existsSync( resolvedDir ) || !fs.statSync( resolvedDir ).isDirectory() ) {
  console.error( `Error: "${ resolvedDir }" is not a valid directory.` );
  process.exit( 1 );
}


// ── Helpers ─────────────────────────────────────────────────────────

/**
 * Normalize a song name for fuzzy matching.
 * Mirror of the function in songs-db.ts.
 *
 * @param name - raw song display name
 * @returns normalized string
 */
function normalizeSongName( name: string ): string {
  let n = name;
  n = n.replace( /\.(pptx?|ppt)$/i, '' );
  n = n.replace( /\s*-\s*Shortcut$/i, '' );
  n = n.replace( /^\d+\s*[-–—]?\s*/, '' );
  n = n.toLowerCase();
  n = n.replace( /[-–—]/g, ' ' );
  n = n.replace( /[^a-z0-9\s]/g, '' );
  n = n.replace( /\s+/g, ' ' ).trim();
  return n;
}


/**
 * Resolve the target of a .lnk shortcut.
 *
 * @param shortcutPath - absolute path to the .lnk file
 * @returns target path or null
 */
function getShortcutTarget( shortcutPath: string ): Promise<string | null> {
  return new Promise( ( resolve ) => {
    ws.query( shortcutPath, ( error: string | null, options?: ws.ShortcutOptions ) => {
      if ( error || !options?.target ) {
        resolve( null );
      } else {
        resolve( options.target );
      }
    });
  });
}


/**
 * Resolve a target path that may contain Windows environment variables.
 *
 * @param p - path possibly containing %VAR% tokens
 * @returns resolved absolute path
 */
function resolveEnvPath( p: string ): string {
  return p.replace( /%([^%]+)%/g, ( _, varName ) => process.env[ varName ] || _ );
}


/**
 * Extract song info (name, number, book) from an absolute target path.
 *
 * @param target - absolute path to the song file
 * @returns { name, number, book }
 */
function extractSongInfo( target: string ): { name: string; number?: string; book?: string } {
  const resolved = resolveEnvPath( target );
  const filename = path.basename( resolved );

  // Extract song number: 3 digits at start
  const numMatch = filename.match( /^(\d{3})\s*-?\s*/ );
  const number = numMatch ? numMatch[ 1 ] : undefined;

  // Clean name: strip number prefix and extension
  const name = filename
    .replace( /^(\d{3})\s*-?\s*/, '' )
    .replace( /\.(pptx?|ppt)$/i, '' )
    .trim();

  // Book is the parent folder
  const parts = resolved.replace( /\\/g, '/' ).split( '/' );
  const book = parts.length >= 2 ? parts[ parts.length - 2 ] : undefined;

  return { name, number, book };
}


/**
 * Escape a string for use in a SQL single-quoted literal.
 *
 * @param s - raw string
 * @returns escaped string
 */
function sqlEscape( s: string ): string {
  return s.replace( /\\/g, '\\\\' ).replace( /'/g, "''" );
}


// ── Shortcut filename parsing ───────────────────────────────────────

interface ParsedShortcut {
  slotType: 'song' | 'chorus' | 'closing';
  slotNumber: number;
}


/**
 * Parse a shortcut filename to determine slot type and number.
 *
 * @param filename - the .lnk filename
 * @returns parsed slot info, or null if not a song/chorus/closing shortcut
 */
function parseShortcutSlot( filename: string ): ParsedShortcut | null {
  const lower = filename.toLowerCase();
  if ( !lower.endsWith( '.lnk' ) ) return null;

  // Closing song
  if ( lower.startsWith( 'closing song ' ) ) {
    return { slotType: 'closing', slotNumber: 1 };
  }

  // Song: "song N ..."
  const songMatch = lower.match( /^song\s+(\d+)\s+/ );
  if ( songMatch ) {
    return { slotType: 'song', slotNumber: parseInt( songMatch[ 1 ] ) };
  }

  // Chorus: "chorus N ..."
  const chorusMatch = lower.match( /^chorus\s+(\d+)\s+/ );
  if ( chorusMatch ) {
    return { slotType: 'chorus', slotNumber: parseInt( chorusMatch[ 1 ] ) };
  }

  return null;
}


// ── Main ────────────────────────────────────────────────────────────

interface SongRecord {
  name: string;
  normalizedName: string;
  number?: string;
  book?: string;
}

interface SelectionRecord {
  sundayDate: string;
  slotType: 'song' | 'chorus' | 'closing';
  slotNumber: number;
  songKey: string;
  displayName: string;
}


/**
 * Main entry point: scan folders, resolve shortcuts, write SQL.
 */
async function main() {
  console.log( `Scanning: ${ resolvedDir }` );

  // Collect unique songs and all selections
  const songMap = new Map<string, SongRecord>();
  const selections: SelectionRecord[] = [];

  // Get all YYYYMMDD folders
  const entries = await fs.promises.readdir( resolvedDir, { withFileTypes: true } );
  const sundayFolders = entries
    .filter( e => e.isDirectory() && /^\d{8}$/.test( e.name ) )
    .map( e => e.name )
    .sort();

  console.log( `Found ${ sundayFolders.length } Sunday folders\n` );

  for ( const folder of sundayFolders ) {
    const folderPath = path.join( resolvedDir, folder );
    let files: string[];
    try {
      files = await fs.promises.readdir( folderPath );
    } catch {
      continue;
    }

    // Separate prefixed shortcuts from bare (old-style) shortcuts
    const allLnk = files.filter( f => f.toLowerCase().endsWith( '.lnk' ) ).sort();
    const prefixed: string[] = [];
    const bare: string[] = [];

    for ( const f of allLnk ) {
      const lower = f.toLowerCase();
      if (
        lower.startsWith( 'song ' ) ||
        lower.startsWith( 'chorus ' ) ||
        lower.startsWith( 'closing song ' )
      ) {
        prefixed.push( f );
      } else {
        bare.push( f );
      }
    }

    /**
     * Process a single shortcut: resolve target, register song, record selection.
     *
     * @param shortcutFile - filename of the .lnk file
     * @param slot - parsed slot type and number
     */
    const processShortcut = async ( shortcutFile: string, slot: ParsedShortcut ) => {
      const shortcutPath = path.join( folderPath, shortcutFile );
      const target = await getShortcutTarget( shortcutPath );

      if ( !target ) {
        console.warn( `  ⚠ Could not resolve: ${ folder }/${ shortcutFile }` );
        return;
      }

      const info = extractSongInfo( target );
      if ( !info.name ) {
        console.warn( `  ⚠ Empty name for: ${ folder }/${ shortcutFile }` );
        return;
      }

      const normalized = normalizeSongName( info.name );
      const songKey = `${ normalized }|${ info.number || '' }|${ ( info.book || '' ).toLowerCase() }`;

      // Register unique song
      if ( !songMap.has( songKey ) ) {
        songMap.set( songKey, {
          name: info.name,
          normalizedName: normalized,
          number: info.number,
          book: info.book,
        });
      }

      const display = info.number ? `${ info.number } - ${ info.name }` : info.name;

      selections.push({
        sundayDate: slot.slotType === 'closing' ? folder.slice( 0, 6 ) : folder,
        slotType: slot.slotType,
        slotNumber: slot.slotNumber,
        songKey,
        displayName: display,
      });

      console.log( `  ${ folder } → ${ slot.slotType } ${ slot.slotNumber }: ${ display }` );
    };

    // Process prefixed shortcuts (modern format)
    for ( const shortcutFile of prefixed ) {
      const slot = parseShortcutSlot( shortcutFile );
      if ( !slot ) continue;
      await processShortcut( shortcutFile, slot );
    }

    // Process bare shortcuts (old format: "115 - Ivory Palaces - Shortcut.lnk")
    // Assign sequential song slots starting at 1
    if ( bare.length > 0 && prefixed.length === 0 ) {
      let songSlot = 1;
      for ( const shortcutFile of bare ) {
        await processShortcut( shortcutFile, { slotType: 'song', slotNumber: songSlot } );
        songSlot++;
      }
    }
  }

  console.log( `\nUnique songs: ${ songMap.size }` );
  console.log( `Total selections: ${ selections.length }` );

  // ── Generate SQL ──────────────────────────────────────────────

  const lines: string[] = [];
  lines.push( '-- ============================================================' );
  lines.push( '-- Song Library & Selection History Import' );
  lines.push( `-- Generated: ${ new Date().toISOString() }` );
  lines.push( `-- Source: ${ resolvedDir }` );
  lines.push( `-- Folders scanned: ${ sundayFolders.length }` );
  lines.push( `-- Unique songs: ${ songMap.size }` );
  lines.push( `-- Total selections: ${ selections.length }` );
  lines.push( '-- ============================================================' );
  lines.push( '' );
  lines.push( '' );

  // ── Part 1: Song inserts ──────────────────────────────────────

  lines.push( '-- ============================================================' );
  lines.push( '-- SONGS' );
  lines.push( '-- ============================================================' );
  lines.push( '' );

  // Sort songs by book, then number, then name
  const sortedSongs = Array.from( songMap.entries() ).sort( ( [ , a ], [ , b ] ) => {
    const bookCmp = ( a.book || '' ).localeCompare( b.book || '' );
    if ( bookCmp !== 0 ) return bookCmp;
    if ( a.number && b.number ) return a.number.localeCompare( b.number, undefined, { numeric: true } );
    if ( a.number ) return -1;
    if ( b.number ) return 1;
    return a.name.localeCompare( b.name );
  });

  let currentBook = '';

  for ( const [ , song ] of sortedSongs ) {
    // Book section comment
    const book = song.book || 'Unknown';
    if ( book !== currentBook ) {
      if ( currentBook ) lines.push( '' );
      lines.push( `-- ${ book }` );
      lines.push( `-- ${ '-'.repeat( book.length ) }` );
      currentBook = book;
    }

    // Song comment
    const display = song.number ? `${ song.number } - ${ song.name }` : song.name;
    lines.push( `-- ${ display }` );

    // Build INSERT
    const nameEsc = sqlEscape( song.name );
    const normalizedEsc = sqlEscape( song.normalizedName );
    const numberVal = song.number ? `'${ sqlEscape( song.number ) }'` : 'NULL';
    const bookVal = song.book ? `'${ sqlEscape( song.book ) }'` : 'NULL';

    lines.push(
      `INSERT INTO songs ( name, normalized_name, number, book ) ` +
      `VALUES ( '${ nameEsc }', '${ normalizedEsc }', ${ numberVal }, ${ bookVal } ) ` +
      `ON DUPLICATE KEY UPDATE ` +
      `number = COALESCE( songs.number, VALUES( number ) ), ` +
      `book = COALESCE( songs.book, VALUES( book ) );`
    );
  }

  lines.push( '' );
  lines.push( '' );

  // ── Part 2: Selection inserts by Sunday ───────────────────────

  lines.push( '-- ============================================================' );
  lines.push( '-- SONG SELECTIONS' );
  lines.push( '-- ============================================================' );
  lines.push( '' );

  // Group selections by sundayDate
  const selectionsByDate = new Map<string, SelectionRecord[]>();
  for ( const sel of selections ) {
    const existing = selectionsByDate.get( sel.sundayDate ) || [];
    existing.push( sel );
    selectionsByDate.set( sel.sundayDate, existing );
  }

  // Sort date keys
  const sortedDates = Array.from( selectionsByDate.keys() ).sort();

  for ( const dateKey of sortedDates ) {
    const dateSelections = selectionsByDate.get( dateKey )!;

    // Format date for comment
    let dateLabel = dateKey;
    if ( dateKey.length === 8 ) {
      const y = dateKey.slice( 0, 4 );
      const m = parseInt( dateKey.slice( 4, 6 ) );
      const d = parseInt( dateKey.slice( 6, 8 ) );
      const months = [ '', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec' ];
      dateLabel = `${ months[ m ] } ${ d }, ${ y }`;
    } else if ( dateKey.length === 6 ) {
      const y = dateKey.slice( 0, 4 );
      const m = parseInt( dateKey.slice( 4, 6 ) );
      const months = [ '', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec' ];
      dateLabel = `${ months[ m ] } ${ y } (closing)`;
    }

    lines.push( `-- ${ dateLabel }` );

    // Sort: songs first (by slot number), then choruses, then closing
    const slotOrder: Record<string, number> = { song: 1, chorus: 2, closing: 3 };
    dateSelections.sort( ( a, b ) => {
      const typeOrd = ( slotOrder[ a.slotType ] || 9 ) - ( slotOrder[ b.slotType ] || 9 );
      if ( typeOrd !== 0 ) return typeOrd;
      return a.slotNumber - b.slotNumber;
    });

    for ( const sel of dateSelections ) {
      const song = songMap.get( sel.songKey )!;
      const normalizedEsc = sqlEscape( song.normalizedName );
      const numberVal = song.number ? `'${ sqlEscape( song.number ) }'` : 'NULL';
      const bookVal = song.book ? `'${ sqlEscape( song.book ) }'` : 'NULL';

      // Comment for this selection
      const slotLabel = sel.slotType === 'closing' ? 'Closing Song' :
        `${ sel.slotType.charAt( 0 ).toUpperCase() + sel.slotType.slice( 1 ) } ${ sel.slotNumber }`;
      lines.push( `--   ${ slotLabel }: ${ sel.displayName }` );

      // Use a subquery to find the song_id by normalized_name + number + book
      let whereClause = `normalized_name = '${ normalizedEsc }'`;
      if ( song.number ) {
        whereClause += ` AND number = ${ numberVal }`;
      }
      if ( song.book ) {
        whereClause += ` AND book = ${ bookVal }`;
      }

      lines.push(
        `INSERT INTO song_selections ( song_id, sunday_date, slot_type, slot_number ) ` +
        `SELECT id, '${ sqlEscape( sel.sundayDate ) }', '${ sel.slotType }', ${ sel.slotNumber } ` +
        `FROM songs WHERE ${ whereClause } LIMIT 1;`
      );
    }

    lines.push( '' );
  }

  // ── Write output into the scanned folder ──────────────────────

  const outputFile = path.join( resolvedDir, 'song-inserts.sql' );
  await fs.promises.writeFile( outputFile, lines.join( '\n' ), 'utf8' );

  console.log( `\n✅ SQL written to: ${ outputFile }` );
}


main().catch( err => {
  console.error( 'Fatal error:', err );
  process.exit( 1 );
});
