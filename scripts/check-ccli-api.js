const http = require( 'http' );

http.get( 'http://localhost:8080/api/songs', ( res ) => {
  let data = '';
  res.on( 'data', chunk => data += chunk );
  res.on( 'end', () => {
    const songs = JSON.parse( data );
    const withCcli = songs.filter( s => s.ccli );
    const withNull = songs.filter( s => s.ccli === null );
    const withUndef = songs.filter( s => s.ccli === undefined );
    const noProp = songs.filter( s => !( 'ccli' in s ) );

    console.log( 'Total songs:', songs.length );
    console.log( 'With ccli value:', withCcli.length );
    console.log( 'ccli === null:', withNull.length );
    console.log( 'ccli === undefined:', withUndef.length );
    console.log( 'No ccli property:', noProp.length );
    console.log( '\nSample without ccli:' );
    songs.filter( s => !s.ccli ).slice( 0, 3 ).forEach( s => {
      console.log( `  ${ s.name } → ccli: ${ JSON.stringify( s.ccli ) }, has prop: ${ 'ccli' in s }` );
    });
  });
});
