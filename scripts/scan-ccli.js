/**
 * Scan all pptx songs for CCLI numbers on slide 1 and populate
 * the songs table where ccli is currently NULL.
 */
const fs = require( 'fs' );
const path = require( 'path' );
const JSZip = require( 'jszip' );
const mysql = require( 'mysql2/promise' );
require( 'dotenv' ).config();

/**
 * Extract CCLI song number from the first slide of a pptx file.
 *
 * @param filePath - absolute path to the pptx file
 * @returns CCLI number string, or null if not found
 */
async function extractCcli( filePath ) {
  try {
    const data = fs.readFileSync( filePath );
    const zip = await JSZip.loadAsync( data );

    const slide1 = zip.file( 'ppt/slides/slide1.xml' );
    if ( !slide1 ) return null;

    const xml = await slide1.async( 'string' );

    // Extract all text runs in order
    const textRuns = [];
    const re = /<a:t>([^<]*)<\/a:t>/g;
    let match;
    while ( ( match = re.exec( xml ) ) !== null ) {
      textRuns.push( match[ 1 ] );
    }

    // Join all text and search for CCLI Song # pattern
    const fullText = textRuns.join( '' );
    const ccliMatch = fullText.match( /CCLI\s*Song\s*#\s*(\d{4,8})/i );
    if ( ccliMatch ) return ccliMatch[ 1 ];

    // Also try "CCLI #" followed by digits (less specific)
    const altMatch = fullText.match( /CCLI\s*#\s*(\d{4,8})/i );
    if ( altMatch ) return altMatch[ 1 ];

    return null;
  } catch {
    return null;
  }
}

( async () => {
  const pool = mysql.createPool({
    host: process.env.MYSQL_HOST,
    port: parseInt( process.env.MYSQL_PORT || '3306' ),
    user: process.env.MYSQL_USER,
    password: process.env.MYSQL_PASSWORD,
    database: process.env.MYSQL_DATABASE,
  });

  const songsDir = process.env.SONGS_DIRECTORY;

  // Get songs with no CCLI that have a file_path
  const [ songs ] = await pool.query(
    'SELECT id, name, number, book, file_path FROM songs WHERE ccli IS NULL AND file_path IS NOT NULL'
  );

  console.log( `Scanning ${ songs.length } songs with no CCLI...` );

  let found = 0;
  let scanned = 0;
  let errors = 0;

  for ( const song of songs ) {
    const filePath = path.join( songsDir, song.file_path );
    if ( !fs.existsSync( filePath ) ) continue;

    scanned++;
    const ccli = await extractCcli( filePath );

    if ( ccli ) {
      await pool.query( 'UPDATE songs SET ccli = ? WHERE id = ?', [ ccli, song.id ] );
      found++;
      console.log( `  ${ song.number ? song.number + ' - ' : '' }${ song.name } → CCLI ${ ccli }` );
    }
  }

  // Also scan songs without file_path by walking the songs directory
  const [ noPath ] = await pool.query(
    'SELECT id, name, number, book, normalized_name FROM songs WHERE ccli IS NULL AND file_path IS NULL'
  );

  if ( noPath.length > 0 ) {
    console.log( `\nTrying to find files for ${ noPath.length } songs without file_path...` );

    // Build a lookup of all pptx files
    const fileMap = new Map();
    const walk = ( dir ) => {
      for ( const entry of fs.readdirSync( dir, { withFileTypes: true } ) ) {
        const full = path.join( dir, entry.name );
        if ( entry.isDirectory() ) {
          walk( full );
        } else if ( entry.name.toLowerCase().endsWith( '.pptx' ) ) {
          // Normalize the filename for matching
          let normalized = entry.name.replace( /\.pptx$/i, '' );
          normalized = normalized.replace( /^\d+\s*[-–—]?\s*/, '' ).toLowerCase().replace( /[^a-z0-9\s]/g, '' ).replace( /\s+/g, ' ' ).trim();
          fileMap.set( normalized, { full, relative: path.relative( songsDir, full ) } );
        }
      }
    };
    walk( songsDir );

    for ( const song of noPath ) {
      const fileInfo = fileMap.get( song.normalized_name );
      if ( !fileInfo ) continue;

      scanned++;
      const ccli = await extractCcli( fileInfo.full );

      // Update file_path too since we found it
      if ( ccli ) {
        await pool.query( 'UPDATE songs SET ccli = ?, file_path = ? WHERE id = ?', [ ccli, fileInfo.relative, song.id ] );
        found++;
        console.log( `  ${ song.number ? song.number + ' - ' : '' }${ song.name } → CCLI ${ ccli } (+ file_path)` );
      } else {
        await pool.query( 'UPDATE songs SET file_path = ? WHERE id = ?', [ fileInfo.relative, song.id ] );
      }
    }
  }

  console.log( `\nDone. Scanned: ${ scanned }, CCLI found: ${ found }` );
  await pool.end();
})();
