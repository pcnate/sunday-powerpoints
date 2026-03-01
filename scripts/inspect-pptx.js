/**
 * Inspect first slide of pptx files to find CCLI number patterns.
 */
const fs = require( 'fs' );
const path = require( 'path' );
const JSZip = require( 'jszip' );
require( 'dotenv' ).config();

const songsDir = process.env.SONGS_DIRECTORY;
const files = [
  path.join( songsDir, 'Hymnal', '001 Everybody Sing Praise to the Lord.pptx' ),
  path.join( songsDir, 'Hymnal', '016 How Great Thou Art.pptx' ),
  path.join( songsDir, 'Hymnal', '054 Great is Thy Faithfulness.pptx' ),
  path.join( songsDir, 'Choruses', 'Because He Lives.pptx' ),
  path.join( songsDir, 'Choruses', 'Lord I Need You.pptx' ),
];

( async () => {
  for ( const file of files ) {
    if ( !fs.existsSync( file ) ) {
      console.log( `SKIP (not found): ${ path.basename( file ) }` );
      continue;
    }

    console.log( `\n=== ${ path.basename( file ) } ===` );

    const data = fs.readFileSync( file );
    const zip = await JSZip.loadAsync( data );

    // Read slide1.xml
    const slide1 = zip.file( 'ppt/slides/slide1.xml' );
    if ( !slide1 ) {
      console.log( '  No slide1.xml found' );
      continue;
    }

    const xml = await slide1.async( 'string' );

    // Extract all text runs
    const textRuns = [];
    const re = /<a:t>([^<]*)<\/a:t>/g;
    let match;
    while ( ( match = re.exec( xml ) ) !== null ) {
      if ( match[ 1 ].trim() ) textRuns.push( match[ 1 ] );
    }

    console.log( 'Text on slide 1:' );
    textRuns.forEach( t => console.log( `  "${ t }"` ) );

    // Look for CCLI-like patterns
    const fullText = textRuns.join( ' ' );
    const ccliMatch = fullText.match( /CCLI\s*(?:#|Song\s*#?:?\s*)?\s*(\d{4,8})/i );
    if ( ccliMatch ) {
      console.log( `  >>> CCLI FOUND: ${ ccliMatch[ 1 ] }` );
    } else {
      console.log( '  >>> No CCLI pattern found' );
    }
  }
})();
