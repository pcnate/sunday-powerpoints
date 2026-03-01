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

  // Find all duplicate groups by normalized_name + number
  const [ dupes ] = await pool.query(
    'SELECT normalized_name, number, MIN(id) as keep_id, GROUP_CONCAT(id ORDER BY id) as all_ids, COUNT(*) as cnt FROM songs GROUP BY normalized_name, number HAVING cnt > 1'
  );

  let mergedCount = 0;
  let deletedCount = 0;

  for ( const dupe of dupes ) {
    const keepId = dupe.keep_id;
    const allIds = dupe.all_ids.split( ',' ).map( Number );
    const dupeIds = allIds.filter( id => id !== keepId );

    // Reassign all song_selections from duplicates to the keeper
    for ( const dupeId of dupeIds ) {
      const [ result ] = await pool.query(
        'UPDATE song_selections SET song_id = ? WHERE song_id = ?',
        [ keepId, dupeId ]
      );
      mergedCount += result.affectedRows;
    }

    // Delete duplicate song records
    if ( dupeIds.length > 0 ) {
      const placeholders = dupeIds.map( () => '?' ).join( ',' );
      await pool.query(
        `DELETE FROM songs WHERE id IN (${ placeholders })`,
        dupeIds
      );
      deletedCount += dupeIds.length;
    }
  }

  console.log( 'Duplicate groups processed:', dupes.length );
  console.log( 'Selections reassigned:', mergedCount );
  console.log( 'Duplicate songs deleted:', deletedCount );

  // Verify
  const [ remaining ] = await pool.query( 'SELECT COUNT(*) as total FROM songs' );
  const [ sel ] = await pool.query( 'SELECT COUNT(*) as total FROM song_selections WHERE removed_at IS NULL' );
  console.log( 'Songs remaining:', remaining[ 0 ].total );
  console.log( 'Active selections:', sel[ 0 ].total );

  await pool.end();
})();
