import { normalizeSongName, parseShortcutFilename } from '../src/songs-db';


describe( 'normalizeSongName', () => {

  it( 'strips .pptx extension', () => {
    expect( normalizeSongName( 'Amazing Grace.pptx' ) ).toBe( 'amazing grace' );
  } );


  it( 'strips .ppt extension', () => {
    expect( normalizeSongName( 'Amazing Grace.ppt' ) ).toBe( 'amazing grace' );
  } );


  it( 'strips .PPTX extension (case-insensitive)', () => {
    expect( normalizeSongName( 'Amazing Grace.PPTX' ) ).toBe( 'amazing grace' );
  } );


  it( 'strips " - Shortcut" suffix', () => {
    expect( normalizeSongName( 'Amazing Grace - Shortcut' ) ).toBe( 'amazing grace' );
  } );


  it( 'strips leading number prefix with dash', () => {
    expect( normalizeSongName( '203 - Tell Me the Story' ) ).toBe( 'tell me the story' );
  } );


  it( 'strips leading number prefix without dash', () => {
    expect( normalizeSongName( '203 Tell Me the Story' ) ).toBe( 'tell me the story' );
  } );


  it( 'strips leading number prefix with em dash', () => {
    expect( normalizeSongName( '203\u2014Tell Me the Story' ) ).toBe( 'tell me the story' );
  } );


  it( 'replaces hyphens with spaces', () => {
    expect( normalizeSongName( 'All-In-All' ) ).toBe( 'all in all' );
  } );


  it( 'removes punctuation', () => {
    expect( normalizeSongName( "How Great Thou Art!" ) ).toBe( 'how great thou art' );
  } );


  it( 'removes apostrophes and commas', () => {
    expect( normalizeSongName( "It's Well, With My Soul" ) ).toBe( 'its well with my soul' );
  } );


  it( 'collapses whitespace', () => {
    expect( normalizeSongName( '  Amazing   Grace  ' ) ).toBe( 'amazing grace' );
  } );


  it( 'handles combined extension + number + suffix', () => {
    expect( normalizeSongName( '203 - Tell Me the Story - Shortcut.pptx' ) ).toBe( 'tell me the story' );
  } );


  it( 'preserves digits in song names', () => {
    expect( normalizeSongName( 'Psalm 23' ) ).toBe( 'psalm 23' );
  } );


  it( 'returns empty string for empty input', () => {
    expect( normalizeSongName( '' ) ).toBe( '' );
  } );


  it( 'returns empty string for number-only input', () => {
    expect( normalizeSongName( '203' ) ).toBe( '' );
  } );
} );


describe( 'parseShortcutFilename', () => {

  describe( 'song shortcuts', () => {

    it( 'parses song with number and name', () => {
      expect( parseShortcutFilename( 'song 1 203 - Tell Me the Story - Shortcut.lnk' ) ).toEqual({
        slotType: 'song',
        slotNumber: 1,
        slotSuffix: undefined,
        songName: 'Tell Me the Story',
        songNumber: '203',
      });
    } );


    it( 'parses song without number', () => {
      expect( parseShortcutFilename( 'song 2 Amazing Grace - Shortcut.lnk' ) ).toEqual({
        slotType: 'song',
        slotNumber: 2,
        slotSuffix: undefined,
        songName: 'Amazing Grace',
        songNumber: undefined,
      });
    } );


    it( 'parses split slot with suffix', () => {
      expect( parseShortcutFilename( 'song 2a 101 - Holy Holy Holy - Shortcut.lnk' ) ).toEqual({
        slotType: 'song',
        slotNumber: 2,
        slotSuffix: 'a',
        songName: 'Holy Holy Holy',
        songNumber: '101',
      });
    } );


    it( 'parses split slot b without number', () => {
      expect( parseShortcutFilename( 'song 2b Great Is Thy Faithfulness - Shortcut.lnk' ) ).toEqual({
        slotType: 'song',
        slotNumber: 2,
        slotSuffix: 'b',
        songName: 'Great Is Thy Faithfulness',
        songNumber: undefined,
      });
    } );
  } );


  describe( 'chorus shortcuts', () => {

    it( 'parses chorus with number', () => {
      expect( parseShortcutFilename( 'chorus 1 045 - Blessed Assurance - Shortcut.lnk' ) ).toEqual({
        slotType: 'chorus',
        slotNumber: 1,
        songName: 'Blessed Assurance',
        songNumber: '045',
      });
    } );


    it( 'parses chorus without number', () => {
      expect( parseShortcutFilename( 'chorus 3 How Great Thou Art - Shortcut.lnk' ) ).toEqual({
        slotType: 'chorus',
        slotNumber: 3,
        songName: 'How Great Thou Art',
        songNumber: undefined,
      });
    } );
  } );


  describe( 'closing song shortcuts', () => {

    it( 'parses closing song with number', () => {
      expect( parseShortcutFilename( 'closing song 203 - Tell Me the Story - Shortcut.lnk' ) ).toEqual({
        slotType: 'closing',
        slotNumber: 1,
        songName: 'Tell Me the Story',
        songNumber: '203',
      });
    } );


    it( 'parses closing song without number', () => {
      expect( parseShortcutFilename( 'closing song Amazing Grace - Shortcut.lnk' ) ).toEqual({
        slotType: 'closing',
        slotNumber: 1,
        songName: 'Amazing Grace',
        songNumber: undefined,
      });
    } );
  } );


  describe( 'invalid filenames', () => {

    it( 'returns null for non-.lnk files', () => {
      expect( parseShortcutFilename( 'song 1 Test.pptx' ) ).toBeNull();
    } );


    it( 'returns null for unrecognized pattern', () => {
      expect( parseShortcutFilename( 'notes.lnk' ) ).toBeNull();
    } );


    it( 'returns null for empty string', () => {
      expect( parseShortcutFilename( '' ) ).toBeNull();
    } );
  } );
} );
