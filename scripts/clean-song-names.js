/**
 * Clean song names in the songs table:
 * - Strip file extensions (.pptx, .ppt)
 * - Strip leading number prefixes (e.g., "217 " or "217 - ")
 *   since number is stored separately
 */
const mysql = require( 'mysql2/promise' );
require( 'dotenv' ).config();

( async () => {
  const pool = mysql.createPool({
    host: process.env.MYSQL_HOST,
    port: parseInt( process.env.MYSQL_PORT || '3306' ),
    user: process.env.MYSQL_USER,
    password: process.env.MYSQL_PASSWORD,
    database: process.env.MYSQL_DATABASE,
  });

  const [ songs ] = await pool.query( 'SELECT id, name, number FROM songs' );
  let updated = 0;

  for ( const song of songs ) {
    let clean = song.name;

    // Strip file extension
    clean = clean.replace( /\.(pptx?|ppt)$/i, '' );

    // Strip leading number prefix if number column is populated
    if ( song.number ) {
      clean = clean.replace( /^\d+\s*[-–—]?\s*/, '' );
    }

    clean = clean.trim();

    if ( clean !== song.name ) {
      await pool.query( 'UPDATE songs SET name = ? WHERE id = ?', [ clean, song.id ] );
      updated++;
      if ( updated <= 10 ) {
        console.log( `  "${ song.name }" -> "${ clean }"` );
      }
    }
  }

  console.log( `\nUpdated ${ updated } of ${ songs.length } song names` );
  await pool.end();
})();
