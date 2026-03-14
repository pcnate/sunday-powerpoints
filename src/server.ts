import path from 'node:path';
import fs from 'node:fs';
import child_process from 'node:child_process';
import express, { Request, Response } from 'express';
import dotenv from 'dotenv';
import ws from 'windows-shortcuts';
const schedule = require( 'node-schedule' );
import { createShortcut, sundaysInMonth, resolveToAbsolutePath } from './sunday-powerpoints';
import { getClosingSong, setClosingSong, syncSelectionsFromFolder, findOrCreateSong, normalizeSongName, recordSongSelection, removeSongSelection, removeAllChorusSelections, FolderSongInfo } from './songs-db';
import { Server as SocketIOServer } from 'socket.io';
import http from 'http';
import { initDb, closeDb, getPool, healthCheck } from './db';
import { RowDataPacket, ResultSetHeader } from 'mysql2/promise';
import { createJobRoutes, recoverStaleJobs, createJobIfNotExists, createProbeJobIfNotExists, createHeldProbeJob } from './routes/jobs';
import { getVideosByPaths, getOldestMtsVideo, deleteVideo } from './videos-db';
import swaggerUi from 'swagger-ui-express';
import YAML from 'yaml';

dotenv.config();


/**
 * Class representing a cache for Sunday folders
 * 
 * This class provides methods to get, set, and delete SundayFolder objects
 * from a cache based on their date (YYYYMMDD format).
 */
class FolderCache {
  size: number;
  public cache: { [date: string]: SundayFolder };
  private garbage: { [date: string]: SundayFolder };

  constructor() {
    this.size = 0;
    this.cache = {};
    this.garbage = {};
  }


  /**
   * Get a SundayFolder by date
   * 
   * @param date - date in YYYYMMDD format
   * @returns SundayFolder object or undefined if not found
   */
  get( date: string ): SundayFolder | undefined {
    return this.cache[ date ];
  }


  /**
   * Set a SundayFolder in the cache
   * 
   * @param date - date in YYYYMMDD format
   * @param folder - SundayFolder object to set in the cache
   */
  set( date: string, folder: SundayFolder ): void {
    this.cache[ date ] = folder;
    this.size = Object.keys( this.cache ).length;
  }


  /**
   * Check if a SundayFolder exists in the cache
   * 
   * @param date - date in YYYYMMDD format
   * @returns true if the folder exists, false otherwise
   */
  delete( date: string ): void {
    this.garbage[ date ] = this.cache[ date ];
    delete this.cache[ date ];
    this.size = Object.keys( this.cache ).length;

    // remove from garbage after 10 minutes
    setTimeout( () => {
      delete this.garbage[ date ];
    }, 10 * 60 * 1000);
  }
}


/**
 * Class representing a Sunday folder
 * 
 * This class creates a SundayFolder object with properties for the folder name,
 * the path, whether it has a presentation file, thumbnail image, video files,
 * 
 * @property {string} name - Name of the folder in YYYYMMDD format
 * @property {string} path - Absolute path to the folder which ends in YYYYMMDD
 * @property {boolean} presentation - true if a PowerPoint file with the configured extension exists
 * @property {string|null} thumbnail - Path to thumbnail image, or null if not found
 * @property {Video[]} video - Array of video files in the folder
 * @property {Record<string, Song|null>} songs - Song shortcuts keyed by slot ID (e.g., "1", "2a", "2b", "3")
 * @property {Chorus[]} choruses - Array of chorus shortcuts in the folder
 * @property {string|null} hasNotes - path to notes file if it exists, or null if not found
 */
class SundayFolder {
  name: string; // YYYYMMDD
  path: string; // absolute path to the folder which ends in YYYYMMDD
  presentation: string | null; // path to PowerPoint file if it exists, or null if not found
  thumbnail: string | null; // path to thumbnail image, or null if not found
  video: Video[]; // video files
  songs: Record<string, Song | null>;
  choruses: Chorus[];
  hasNotes: string | null;

  archived: boolean; // archived folders are in ./YYYY/YYYYMMDD instead of ./YYYYMMDD
  backlog: boolean; // backlog folders are in ./YYYY Backlog/YYYYMMDD instead of ./YYYYMMDD

  /**
   * Create a SundayFolder instance
   *
   * @param name - name of the folder in YYYYMMDD format
   * @param path - absolute path to the folder
   */
  constructor( name: string, path: string ) {
    this.name = name;
    this.path = path;
    this.presentation = null;
    this.thumbnail = null;
    this.video = [];
    this.songs = { '1': null, '2': null, '3': null };
    this.choruses = [];
    this.hasNotes = null;

    // Determine if the folder is archived or backlog based on its path
    this.archived = false;
    this.backlog = false;
    if ( this.path.includes( 'Backlog' ) ) {
      this.backlog = true;
    } else
    // regex match path
    if ( this.path.match( /\/\d{4}\/\d{4}-?\d{2}-?\d{2}$/ ) ) {
      this.archived = true;
    }

    // Check if the folder has a PowerPoint file with the configured extension
    this.scanFolder()
  }


  /**
   * Check if the folder has a PowerPoint file with the configured extension
   *
   * @param _path - absolute path to the video file
   */
  addVideo( _path: fs.PathLike ): void {
    if ( this.video.some( v => v.path === _path.toString() ) ) return; // already exists
    this.video.push( new Video( _path ) );
  }


  /**
   * Remove a video from the list
   * 
   * @param path - absolute path to the video file
   */
  removeVideo( path: fs.PathLike ): void {
    const index = this.video.findIndex( v => v.path === path.toString() );
    if ( index !== -1 ) this.video.splice( index, 1 );
  }


  /**
   * Set a song for the given slot ID.
   *
   * @param slotId - slot identifier (e.g., "1", "2", "2a", "2b", "3")
   * @param _target - absolute path to the target song file
   */
  async setSong( slotId: string, _target: fs.PathLike ) {
    if ( this.songs[ slotId ] ) await this.songs[ slotId ]!.delete();
    const song = new Song( `song ${ slotId }`, _target );
    this.songs[ slotId ] = song;
  }


  /**
   * Add a chorus to the end of the choruses list
   * 
   * @param _target - absolute path to the target chorus file
   */
  addChorus( _target: fs.PathLike ) {
    // determine what the current chorus number is
    const chorusNumber = this.choruses.length + 1;
    const chorus = new Chorus( `chorus ${ chorusNumber }`, _target );
    this.choruses.push( chorus );
  }


  /**
   * delete a chorus from the array after running its delete method
   * then update the chorus prefix with the new numbers
   */
  deleteChorus( chorus: Chorus ) {
    chorus.delete();
    this.choruses = this.choruses.filter( c => c !== chorus );
    this.choruses.forEach( ( _chorus, index ) => {
      _chorus.changePrefix( `chorus ${ index + 1 }` );
    });
  }


  /**
   * Scan the folder for files and set properties accordingly
   */
  async scanFolder() {
    const files = await fs.promises.readdir( this.path );
    let changes = false;

    const choruses: { name: string, path: string, target: string | null }[] = [];

    for ( const file of files ) {
      const filePath = path.join( this.path, file );
      const stat = await fs.promises.stat( filePath );
      
      if ( stat.isFile() ) {
        // Check for PowerPoint file with configured extension
        if ( file.toLowerCase().endsWith( '.' + config.ext.toLowerCase() ) ) {
          if ( !this.presentation ) {
            this.presentation = filePath;
            changes = true;
          }
          continue; // skip further checks for PowerPoint file
        }

        // Check for notes file
        if ( /^(\d{4}-?\d{2}-?\d{2}[\s-]?)?notes\.txt$/i.test( file ) ) {
          if ( this.hasNotes !== filePath ) {
            this.hasNotes = filePath;
            changes = true;
          }
          continue; // skip further checks for notes file
        }

        // Check for song shortcuts and get the target file
        if ( file.toLowerCase().startsWith( 'song ' ) ) {
          const songMatch = file.match( /^song (\d+)([a-z])? /i );
          if ( !songMatch ) continue; // not a song shortcut

          const slotId = songMatch[ 2 ] ? `${ songMatch[ 1 ] }${ songMatch[ 2 ].toLowerCase() }` : songMatch[ 1 ];
          const target = await getShortcutTarget( filePath );
          const song = new Song( `song ${ slotId }`, target || '', false );

          if ( this.songs[ slotId ] && this.songs[ slotId ]!.target === song.target ) continue; // already exists
          this.songs[ slotId ] = song;
          changes = true;
          continue;
        }

        // Check for chorus shortcuts
        if ( file.toLowerCase().startsWith( 'chorus ' ) ) {
          const target = await getShortcutTarget( filePath );
          choruses.push({ name: file, path: filePath, target });
        }

      }

      // post process the choruses
      if ( choruses.length > 0 ) {
        choruses.sort( ( a, b ) => a.name.localeCompare( b.name ) ).forEach( async chorus => {
          // extract the prefix 
          const prefix = chorus.name.match( /^(chorus\s*\d*)/i )?.[0] || 'chorus';
          const target = chorus.target || '';
          const _chorus = new Chorus( prefix, target, false );
          // check if the chorus already exists
          if ( this.choruses.some( c => c.path === _chorus.path ) ) return; // already exists
          this.choruses.push( _chorus );
          changes = true;
        });
      }
    }

    const videoFiles = await fs.promises.readdir( path.join( this.path, 'Vids' ) ).catch( () => [] );
    for ( const videoFile of videoFiles ) {
      const videoPath = path.join( this.path, 'Vids', videoFile );
      const stat = await fs.promises.stat( videoPath );

      if ( stat.isFile() ) {
        // if the video is a thm file, rename to jpeg and use as thumbnail
        if ( videoFile.toLocaleLowerCase().endsWith( '.thm' ) ) {
          const newPath = videoPath.replace( /\.thm$/i, '.jpeg' );
          await fs.promises.rename( videoPath, newPath );
          this.thumbnail = newPath;
          changes = true;
          continue;
        }

        // jpeg thumbnails — set as thumbnail and skip addVideo
        if ( videoFile.toLowerCase().endsWith( '.jpeg' ) ) {
          if ( this.thumbnail !== videoPath ) {
            this.thumbnail = videoPath;
            changes = true;
          }
          continue;
        }

        this.addVideo( videoPath );
      }
    }

    // TODO: send an event to the socketio client that changes were detected
    if ( changes ) {
      io.emit( 'folder-changes', this.toJSON() );
    }
  }


  /**
   * Convert the SundayFolder object to a JSON representation
   * 
   * @returns JSON object with folder properties
   */
  toJSON(): object {
    return {
      name: this.name,
      path: this.path,
      presentation: this.presentation,
      thumbnail: this.thumbnail,
      video: this.video.map( v => v.toJSON() ),
      choruses: this.choruses.map( c => c.toJSON() ),
      songs: Object.fromEntries(
        Object.entries( this.songs ).map( ([ k, v ]) => [ k, v ? v.toJSON() : null ] )
      ),
      song1: this.songs[ '1' ] ? this.songs[ '1' ]!.toJSON() : null,
      song2: this.songs[ '2' ] ? this.songs[ '2' ]!.toJSON() : null,
      song3: this.songs[ '3' ] ? this.songs[ '3' ]!.toJSON() : null,
      hasNotes: this.hasNotes,
      archived: this.archived,
      backlog: this.backlog
    }
  }
}


/**
 * Get the target of a Windows shortcut file
 * 
 * @param shortcutPath - Path to the shortcut file
 */
async function getShortcutTarget( shortcutPath: fs.PathLike ): Promise<string | null> {
  return new Promise( ( resolve, reject ) => {
    ws.query( shortcutPath.toString(), ( error: string | null, options?: ws.ShortcutOptions ) => {
      if ( error ) {
        console.error( `Error querying shortcut ${ shortcutPath }:`, error );
        resolve( null );
      } else if ( options && options.target ) {
        resolve( options.target );
      } else {
        resolve( null );
      }
    });
  });
}


/** Resolved OneDrive root (computed once after config is loaded). */
let oneDriveRootNormalised = '';


/**
 * Make any absolute path relative to the OneDrive root.
 * Returns the path with forward slashes, e.g. "PowerPoints/songs/203 - Name.pptx".
 * If the path doesn't fall under the OneDrive root, returns it normalised as-is.
 *
 * @param absolutePath - an absolute filesystem path (may contain env vars like %OneDriveConsumer%)
 * @returns path relative to OneDrive root
 */
function toRelativePath( absolutePath: string ): string {
  if ( !oneDriveRootNormalised ) {
    oneDriveRootNormalised = resolveToAbsolutePath( config.rootPath || '%OneDriveConsumer%' )
      .replace( /\\/g, '/' ).replace( /\/$/, '' );
  }
  const resolved = resolveToAbsolutePath( absolutePath ).replace( /\\/g, '/' );
  if ( resolved.toLowerCase().startsWith( oneDriveRootNormalised.toLowerCase() ) ) {
    return resolved.slice( oneDriveRootNormalised.length + 1 );
  }
  return resolved;
}


/** Cache for resolved shortcut targets (path → relative target). Cleared every 5 minutes. */
const shortcutTargetCache = new Map<string, string>();
setInterval( () => shortcutTargetCache.clear(), 5 * 60 * 1000 );


/**
 * Resolve a shortcut's target path and return it relative to the OneDrive root.
 * Results are cached to avoid repeated ws.query calls for the same shortcut.
 *
 * @param shortcutFullPath - absolute path to the .lnk file
 * @param fallbackName - display name if resolution fails
 * @returns relative path string
 */
async function getRelativeTarget( shortcutFullPath: string, fallbackName: string ): Promise<string> {
  const cached = shortcutTargetCache.get( shortcutFullPath );
  if ( cached !== undefined ) return cached;

  const target = await getShortcutTarget( shortcutFullPath );
  if ( !target ) {
    shortcutTargetCache.set( shortcutFullPath, fallbackName );
    return fallbackName;
  }

  const result = toRelativePath( target );
  shortcutTargetCache.set( shortcutFullPath, result );
  return result;
}


/**
 * Class representing a chorus shortcut
 * 
 * @property {string} name - Name of the video file
 * @property {string} camera - Optional camera name if available
 * @property {string} type - Type of the video file (mp4 or mts)
 * @property {string} path - Absolute path to the video file
 */
class Video {
  name: string;
  camera?: string; // optional camera name if available
  type: 'mp4' | 'mts';
  path: string;


  /**
   * Create a Video object
   * 
   * @param path - absolute path to the video file
   */
  constructor( path: fs.PathLike ) {
    this.path = path.toString();
    this.name = this.path.split( '/' ).pop() || '';
    this.type = this.name.toLowerCase().endsWith( '.mp4' ) ? 'mp4' : 'mts';

    this.camera = 'unknown';

    // TODO add rules for camera name extraction from the path or metadata
    // main camera uses /\d{5}\.(MP4|MTS)/ format
    const cameraMatch = this.name.match( /(\d{5})\.(mp4|mts)$/i );
    if ( cameraMatch ) this.camera = 'main';
    else if ( this.name.toLowerCase().includes( 'camera2' ) )
      this.camera = 'camera2';
    else if ( this.name.toLowerCase().includes( 'camera3' ) )
      this.camera = 'camera3';
  }


  /**
   * Convert the Video object to a JSON representation
   * 
   * @returns JSON object with video properties
   */
  toJSON(): object {
    return {
      name: this.name,
      camera: this.camera,
      type: this.type,
      path: this.path
    };
  }
}


/**
 * Class representing a song shortcut
 * 
 * This class creates a song shortcut object with the given prefix and target file path.
 * 
 * @property {string} name - Name of the song file, derived from the target file path
 * @property {string} prefix - Prefix for the shortcut name, e.g. "song 1"
 * @property {string} path - path of the shortcut file
 * @property {string} target - Absolute path to the target file
 * @property {string} number - Song number, if available (3 digits at the start of the filename)
 * @property {string} book - Book/source of the song, derived from the target file path
 * @property {string} ccli - CCLI number of the song, if available
 */
class Song {
  name: string;
  prefix: string;
  path: string;
  target: string;
  number?: string;
  book?: string;
  ccli?: string;


  /**
   * Constructor for Song class
   * Creates a song shortcut object with the given prefix and target file path.
   * 
   * @param prefix - prefix for the shortcut name, e.g. "song 1"
   * @param _target - absolute path to the target file
   * @param create - whether to create the shortcut file (default: true)
   */
  constructor( prefix: string, _target: fs.PathLike, create: boolean = true ) {
    this.prefix = prefix;
    this.target = _target.toString();
    this.name = this.target.split( path.sep ).pop() || '';
    // Extract song number if filename starts with exactly 3 digits, optionally followed by space or dash
    // Example: `123 - Song Name.pptx` or `123- Song Name.pptx` or `123-Name.pptx`
    const match = this.name.match( /^(\d{3})\s*-?\s*/ );
    this.number = match ? match[ 1 ] : undefined;
    // Book/source is the top-level folder (last part of path)
    this.book = this.target.split( path.sep ).slice( -2, -1 )[ 0 ];
    this.path = path.join( config.outputDirectory, prefix + ' ' + this.name );

    // whether to auto create the shortcut file
    if ( create ) {
      // Create the shortcut file if it does not exist
      const exists = fs.promises.stat( this.path )
      .then( () => true )
      .catch( () => false );

      // Create the shortcut path
      if ( !exists ) {
        createShortcut( this.path, prefix + ' ' + this.name, this.target, config.rootPath || '%OneDriveConsumer%' )
          .then( () => console.log( `Created song shortcut: ${ this.path }` ) )
          .catch( ( error ) => console.error( `Failed to create song shortcut ${ this.path }:`, error ) );
      }
    }
  }


  /**
   * delete the song shortcut
   * 
   * @returns - result object containing success status and new path
   */
  async delete(): Promise<boolean> {
    try {
      await fs.promises.unlink( this.path );
      console.log( `Deleted song shortcut: ${ this.path }` );
      return true;
    } catch ( error ) {
      console.error( `Failed to delete song shortcut ${ this.path }:`, error );
      return false;
    }
  }


  /**
   * Convert the Song object to a JSON representation
   * 
   * @returns JSON object with song properties
   */
  toJSON(): object {
    return {
      name: this.name,
      prefix: this.prefix,
      path: this.path,
      target: this.target,
      number: this.number,
      book: this.book,
      ccli: this.ccli
    };
  }
}


/**
 * Class representing a chorus
 * 
 * This class creates a chorus object with the given prefix and target file path.
 * 
 * @property {string} name - Name of the chorus, derived from the target file path
 * @property {string} prefix - Prefix for the chorus name, e.g. "chorus 1"
 * @property {string} path - path of the shortcut file
 * @property {string} target - Absolute path to the target file
 * @property {string} book - Book/source of the chorus, derived from the target file path
 * @property {string} ccli - CCLI number of the song, if available
 */
class Chorus {
  name: string;
  prefix: string;
  path: string;
  target: string;
  book?: string;
  ccli?: string;


  /**
   * Constructor for Chorus class
   * Creates a chorus object with the given prefix and target file path.
   * 
   * @param prefix - prefix for the chorus name, e.g. "chorus"
   * @param _target - absolute path to the target file
   * @param create - whether to create the shortcut file (default: true)
   */
  constructor( prefix: string, _target: fs.PathLike, create: boolean = true ) {
    this.prefix = prefix;
    this.target = _target.toString();
    this.name = this.target.split( path.sep ).pop() || '';

    // Book/source is the top-level folder (last part of path)
    this.book = this.target.split( path.sep ).slice( -2, -1 )[ 0 ];

    // Create the path for the chorus file
    this.path = path.join( config.outputDirectory, prefix + ' ' + this.name );

    // If create is false, do not create the shortcut file
    if ( !create ) return;

    // determine if the file already exists
    const exists = fs.promises.stat( this.path )
      .then( () => true )
      .catch( () => false );

    if ( !exists ) {
      // If the file does not exist, create a shortcut to the target file
      createShortcut( this.path, prefix + ' ' + this.name, this.target, config.rootPath || '%OneDriveConsumer%' )
        .then( () => console.log( `Created chorus shortcut: ${ this.path }` ) )
        .catch( ( error ) => console.error( `Failed to create chorus shortcut ${ this.path }:`, error ) );
    }
  }


  /**
   * Change the prefix of the chorus name and path
   * 
   * @param prefix - new prefix for the chorus
   * @returns - result object containing success status and new path
   */
  async changePrefix( prefix: string ): Promise<{ success: boolean, path: string }> {
    // No change needed, return current path
    if ( prefix === this.prefix ) {
      return { success: true, path: this.path };
    }

    // Change the prefix of the chorus name and path
    const oldName = this.name;
    this.name = prefix + ' ' + oldName;
    this.path = path.join( path.dirname( this.path ), this.name );

    // attempt to rename the existing file
    const success = await fs.promises.rename( this.path, this.path )
      .then( () => true )
      .catch( ( error ) => {
        console.error( `Failed to rename chorus file ${ this.path }:`, error );
        return false;
      });

    return { success, path: this.path };
  }


  /**
   * Delete the chorus file
   * 
   * @returns - promise that resolves when the file is deleted
   */
  async delete(): Promise<boolean> {
    try {
      await fs.promises.unlink( this.path );
      console.log( `Deleted chorus file: ${ this.path }` );
      return true;
    } catch ( error ) {
      console.error( `Failed to delete chorus file ${ this.path }:`, error );
      return false;
    }
  }


  /**
   * Convert the Chorus object to a JSON representation
   * 
   * @returns JSON object with chorus properties
   */
  toJSON(): object {
    return {
      name: this.name,
      prefix: this.prefix,
      path: this.path,
      target: this.target,
      book: this.book,
      ccli: this.ccli
    };
  }
}


const sundayFolders: FolderCache = new FolderCache();
const videoFolders: string[] = [];

// Read options from environment variables, with defaults as described in roadmap.md
const config = {
  templateFile: process.env.TEMPLATE_FILE, // required
  ext: process.env.EXT || 'pptx',
  templateDirectory: process.env.TEMPLATE_DIRECTORY || '/input',
  outputDirectory: process.env.OUTPUT_DIRECTORY || '/output',
  rootPath: process.env.ROOT_PATH || '%OneDriveConsumer%',
  songsDirectory: process.env.SONGS_DIRECTORY || '/songs', // default to /songs if not set
  songTemplate: process.env.SONG_TEMPLATE || '000 - Song Template.pptx',
  appName: process.env.APP_NAME || 'Sunday PowerPoints',
};

if ( !config.templateFile ) {
  console.error( 'TEMPLATE_FILE environment variable is required.' );
  process.exit( 1 );
}

/**
 * Convert an absolute path to a relative job path by stripping the OneDrive root prefix.
 * Paths are stored relative to %OneDriveConsumer% so workers on any machine
 * can resolve them using their own OneDrive environment variable.
 *
 * @param absolutePath - full filesystem path
 * @returns path relative to OneDrive root (e.g. "PowerPoints/Sermon Videos/Sermons/20250420/Vids/file.mp4")
 */
function toJobPath( absolutePath: string ): string {
  const onedriveRoot = resolveToAbsolutePath( config.rootPath || '%OneDriveConsumer%' )
    .replace( /\\/g, '/' ).replace( /\/+$/, '' );
  const normalized = absolutePath.replace( /\\/g, '/' );
  if ( normalized.startsWith( onedriveRoot + '/' ) ) {
    return normalized.slice( onedriveRoot.length + 1 );
  }
  return normalized;
}

console.log( 'Resolved server config:', JSON.stringify( config, null, 2 ) );

// Add a script to run server.ts in a development mode
if ( process.env.NODE_ENV === 'development' ) {
  console.log( 'Server running in development mode.' );
  // Add any dev-specific logic here
} else {
  console.log( 'Server running in production mode.' );
}

// Start Express web server with HTTP server for socket.io
const app = express();
app.use( express.json() );

const server = http.createServer( app );
const io = new SocketIOServer( server, {
  cors: {
    origin: '*', // Allow all origins (adjust as needed for production)
    methods: [ 'GET', 'POST' ]
  }
});

// Socket.io connection handler
io.on( 'connection', ( socket ) => {
  console.log( 'Socket.io client connected:', socket.id );
  socket.on( 'disconnect', () => {
    console.log( 'Socket.io client disconnected:', socket.id );
  });
});

// Swagger UI — inject version from package.json
const pkgVersion = JSON.parse( fs.readFileSync( path.join( __dirname, '..', 'package.json' ), 'utf-8' ) ).version;
const openapiSpec = YAML.parse( fs.readFileSync( path.join( __dirname, 'openapi.yaml' ), 'utf-8' ) );
openapiSpec.info.version = pkgVersion;
app.use( '/swagger-ui', swaggerUi.serve, swaggerUi.setup( openapiSpec ) );
app.get( '/openapi.yaml', ( _req: Request, res: Response ) => {
  const spec = YAML.parse( fs.readFileSync( path.join( __dirname, 'openapi.yaml' ), 'utf-8' ) );
  spec.info.version = pkgVersion;
  res.type( 'text/yaml' ).send( YAML.stringify( spec ) );
});

// App config endpoint — exposes non-sensitive configuration to the frontend
app.get( '/api/app-config', ( _req: Request, res: Response ) => {
  res.json( { appName: config.appName, version: pkgVersion } );
} );

// Log all /api requests, except /api/webapp-last-modified
app.use( '/api', ( req, res, next ) => {
  if ( req.originalUrl !== '/api/webapp-last-modified' ) {
    console.log( `[${ new Date().toISOString() }] ${ req.method } ${ req.originalUrl }` );
  }
  next();
});

// Mount job API routes
app.use( '/api/jobs', createJobRoutes( io ) );

/**
 * Resolve a YYYYMMDD folder name to its absolute path on disk.
 *
 * Checks three possible locations (in order):
 *   1. OUTPUT_DIRECTORY/YYYYMMDD          (current / upcoming)
 *   2. OUTPUT_DIRECTORY/YYYY/YYYYMMDD     (archived)
 *   3. OUTPUT_DIRECTORY/YYYY Backlog/YYYYMMDD  (backlog)
 *
 * @param folderName - YYYYMMDD string
 * @returns absolute folder path, or null if not found
 */
async function resolveFolderPath( folderName: string ): Promise<string | null> {
  const year = folderName.slice( 0, 4 );
  const candidates = [
    path.join( config.outputDirectory, folderName ),
    path.join( config.outputDirectory, year, folderName ),
    path.join( config.outputDirectory, `${ year } Backlog`, folderName ),
  ];

  for ( const candidate of candidates ) {
    try {
      const stat = await fs.promises.stat( candidate );
      if ( stat.isDirectory() ) return candidate;
    } catch {}
  }

  return null;
}


/**
 * API endpoint to list folders in outputDirectory
 */
/**
 * Scan a single YYYYMMDD folder and return its metadata.
 *
 * @param folderPath - absolute path to the folder
 * @param folderName - the YYYYMMDD folder name
 * @param backlog - whether this folder is in a backlog directory
 * @param archived - whether this folder is in an archived YYYY directory
 * @returns folder metadata object
 */
async function scanFolderMetadata( folderPath: string, folderName: string, backlog: boolean, archived: boolean ) {
  let hasPre = false;
  let isApproved = false;
  let presentationName: string | null = null;
  let presentationSize: string | null = null;
  let presentationModified: string | null = null;
  let files: string[] = [];
  try {
    files = await fs.promises.readdir( folderPath );
    const preFile = files.find( f => f.toLowerCase().endsWith( '.' + config.ext.toLowerCase() ) );
    hasPre = !!preFile;
    // Approved = shortcut in template directory does NOT have "TODO" in its name
    if ( hasPre ) {
      const dateStr = folderName.replace( /(\d{4})(\d{2})(\d{2})/, '$1-$2-$3' );
      const todoLnk = `${ dateStr } TODO.lnk`;
      try {
        await fs.promises.access( path.join( config.templateDirectory, todoLnk ) );
        isApproved = false;
      } catch {
        isApproved = true;
      }
    }

    if ( preFile ) {
      presentationName = preFile;
      try {
        const stat = await fs.promises.stat( path.join( folderPath, preFile ) );
        const kb = stat.size / 1024;
        presentationSize = kb >= 1024 ? `${ ( kb / 1024 ).toFixed( 2 ) } MB` : `${ kb.toFixed( 2 ) } KB`;
        presentationModified = stat.mtime.toLocaleString();
      } catch {}
    }
  } catch {}

  const vidsPath = path.join( folderPath, 'Vids' );
  let hasMp4 = false;
  let hasMkv = false;
  let thumbnail = null;
  let videoFileName = null;
  let hasProduction = false;
  let hasTranscription = false;
  let hasKdenlive = false;
  try {
    const vidsFiles = await fs.promises.readdir( vidsPath );
    hasMp4 = vidsFiles.some( f => f.toLowerCase().endsWith( '.mp4' ) );
    hasMkv = vidsFiles.some( f => f.toLowerCase().endsWith( '.mkv' ) );
    if ( hasMp4 ) {
      const mp4File = vidsFiles.find( f => f.toLowerCase().endsWith( '.mp4' ) );
      videoFileName = mp4File ? toRelativePath( path.join( vidsPath, mp4File ) ) : null;
    }
    // Auto-rename .thm → .jpeg when found
    const thmFile = vidsFiles.find( f => f.toLowerCase().endsWith( '.thm' ) );
    if ( thmFile ) {
      const thmPath = path.join( vidsPath, thmFile );
      const jpegPath = thmPath.replace( /\.thm$/i, '.jpeg' );
      try {
        await fs.promises.rename( thmPath, jpegPath );
        console.log( `[THM] Renamed ${ thmFile } → ${ path.basename( jpegPath ) } in ${ folderName }` );
      } catch {}
    }
    const hasJpeg = vidsFiles.some( f => f.toLowerCase().endsWith( '.jpeg' ) ) || !!thmFile;
    if ( hasJpeg ) {
      thumbnail = `/api/thumbnail/${ encodeURIComponent( folderName ) }`;
    }
    // Pipeline: check for production MP4, transcription VTT, and Kdenlive project in Vids/
    hasProduction = vidsFiles.some( f => /^\d{8}-production\.mp4$/i.test( f ) );
    hasTranscription = vidsFiles.some( f => f.toLowerCase().endsWith( '.vtt' ) );
    hasKdenlive = vidsFiles.some( f => f.toLowerCase().endsWith( '.kdenlive' ) );
  } catch {}

  const hasVideo = hasMp4 || hasMkv;
  const hasThumbnail = !!thumbnail;

  let hasSong1 = false, hasSong2 = false, hasSong3 = false;
  let hasSummary = false;
  const songSlots: string[] = [];
  files.forEach( f => {
    const lower = f.toLowerCase();
    const songMatch = lower.match( /^song (\d+)([a-z])? / );
    if ( songMatch ) {
      const slotId = songMatch[ 2 ] ? `${ songMatch[ 1 ] }${ songMatch[ 2 ] }` : songMatch[ 1 ];
      if ( !songSlots.includes( slotId ) ) songSlots.push( slotId );
      if ( songMatch[ 1 ] === '1' ) hasSong1 = true;
      if ( songMatch[ 1 ] === '2' ) hasSong2 = true;
      if ( songMatch[ 1 ] === '3' ) hasSong3 = true;
    }
    if ( lower.endsWith( '.kdenlive' ) ) hasKdenlive = true;
    if ( /^summary\.txt$/i.test( f ) ) hasSummary = true;
  });

  const notesFile = hasNotesFile( folderName, files );
  const hasNotes = notesFile ? toRelativePath( path.join( folderPath, notesFile ) ) : null;
  return {
    name: folderName,
    hasPre,
    isApproved,
    presentationName,
    presentationSize,
    presentationModified,
    hasMp4,
    hasMkv,
    hasVideo,
    hasThumbnail,
    thumbnail,
    videoFileName,
    hasSong1,
    hasSong2,
    hasSong3,
    songSlots,
    hasNotes,
    hasKdenlive,
    hasProduction,
    hasTranscription,
    hasSummary,
    youtubeUrl: null as string | null,
    sermonPptxFile: null as string | null,
    sermonTitle: null as string | null,
    sermonSpeaker: null as string | null,
    backlog,
    archived,
    files,
  };
}


app.get( '/api/folders', async ( request: Request, response: Response ) => {
  const [ error, entries ]: [ unknown, fs.Dirent[] | null ] = await fs.promises.readdir( config.outputDirectory, { withFileTypes: true } )
    .then( ( entries: fs.Dirent[] ) => [ null, entries ] as [ null, fs.Dirent[] ] )
    .catch( ( error: unknown ) => [ error, null ] as [ unknown, null ] );
  if ( error || !entries ) {
    response.status( 500 ).json( [] );
    return;
  }

  // Collect top-level YYYYMMDD folders, archived YYYY/ subfolders, and backlog subfolders
  const dateFolders: { name: string; fullPath: string; backlog: boolean; archived: boolean }[] = [];
  const dateRegex = /^\d{8}$/;

  for ( const entry of entries ) {
    if ( !entry.isDirectory() ) continue;

    if ( dateRegex.test( entry.name ) ) {
      dateFolders.push({
        name: entry.name,
        fullPath: path.join( config.outputDirectory, entry.name ),
        backlog: false,
        archived: false,
      });
    } else if ( /^\d{4}$/.test( entry.name ) ) {
      // Archived year folder (e.g. "2025")
      const yearPath = path.join( config.outputDirectory, entry.name );
      try {
        const subEntries = await fs.promises.readdir( yearPath, { withFileTypes: true } );
        for ( const sub of subEntries ) {
          if ( sub.isDirectory() && dateRegex.test( sub.name ) ) {
            dateFolders.push({
              name: sub.name,
              fullPath: path.join( yearPath, sub.name ),
              backlog: false,
              archived: true,
            });
          }
        }
      } catch {}
    } else if ( /^\d{4}\s+backlog$/i.test( entry.name ) ) {
      const backlogPath = path.join( config.outputDirectory, entry.name );
      try {
        const subEntries = await fs.promises.readdir( backlogPath, { withFileTypes: true } );
        for ( const sub of subEntries ) {
          if ( sub.isDirectory() && dateRegex.test( sub.name ) ) {
            dateFolders.push({
              name: sub.name,
              fullPath: path.join( backlogPath, sub.name ),
              backlog: true,
              archived: false,
            });
          }
        }
      } catch {}
    }
  }

  dateFolders.sort( ( a, b ) => b.name.localeCompare( a.name ) );

  const folderData = await Promise.all(
    dateFolders.map( f => scanFolderMetadata( f.fullPath, f.name, f.backlog, f.archived ) )
  );

  // Merge YouTube URLs from MySQL (if DB available)
  try {
    const pool = getPool();
    const [ rows ] = await pool.query<RowDataPacket[]>( 'SELECT sunday_date, url FROM youtube_urls' );
    const urlMap = new Map<string, string>();
    for ( const row of rows ) {
      urlMap.set( row.sunday_date, row.url );
    }
    for ( const folder of folderData ) {
      const url = urlMap.get( folder.name );
      if ( url ) folder.youtubeUrl = url;
    }
  } catch {
    // DB not available — youtubeUrl stays null
  }

  // Merge sermon info from MySQL (if DB available)
  try {
    const pool = getPool();
    const [ sermonRows ] = await pool.query<RowDataPacket[]>( 'SELECT sunday_date, pptx_file, title, speaker FROM sermon_info' );
    const sermonMap = new Map<string, { pptxFile: string | null; title: string | null; speaker: string | null }>();
    for ( const row of sermonRows ) {
      sermonMap.set( row.sunday_date, { pptxFile: row.pptx_file, title: row.title, speaker: row.speaker } );
    }
    for ( const folder of folderData ) {
      const info = sermonMap.get( folder.name );
      if ( info ) {
        folder.sermonPptxFile = info.pptxFile;
        folder.sermonTitle = info.title;
        folder.sermonSpeaker = info.speaker;
      }
    }
  } catch {
    // DB not available — sermon info stays null
  }

  response.json( folderData );
});


/**
 * Check for notes file in a folder
 *
 * @param folderName - name of the folder
 * @param files - list of filenames in the folder
 * @returns the notes filename if found, or null
 */
function hasNotesFile( folderName: string, files: string[] ): string | null {
  // Match notes file: {date} Notes.txt (case-insensitive)
  const notesRegex = / Notes\.txt$/i;
  return files.find( f => notesRegex.test( f ) ) || null;
}


/**
 * Endpoint to serve thumbnail images as JPEGs (no query string, auto-detect first .jpeg or .thm)
 */
app.get( '/api/thumbnail/:folder', async ( req: Request, res: Response ) => {
  const folder = req.params.folder;
  if ( !folder ) {
    res.status( 400 ).send( 'Missing folder' );
    return;
  }
  const folderPath = await resolveFolderPath( folder );
  if ( !folderPath ) {
    res.status( 404 ).send( 'Folder not found' );
    return;
  }
  const vidsPath = path.join( folderPath, 'Vids' );
  let files: string[] = [];
  try {
    files = await fs.promises.readdir( vidsPath );
  } catch {
    res.status( 404 ).send( 'Vids folder not found' );
    return;
  }
  const thumbFile = files.find( f => f.toLowerCase().endsWith( '.jpeg' ) || f.toLowerCase().endsWith( '.thm' ) );
  if ( !thumbFile ) {
    res.status( 404 ).send( 'Thumbnail not found' );
    return;
  }
  const filePath = path.join( vidsPath, thumbFile );
  fs.promises.readFile( filePath )
    .then( data => {
      res.setHeader( 'Content-Type', 'image/jpeg' );
      res.send( data );
    })
    .catch( () => {
      res.status( 404 ).send( 'Thumbnail not found' );
    });
});


/**
 * API endpoint to provide template file details
 */
app.get( '/api/template-info', async ( req: Request, res: Response ) => {
  const templateFile = config.templateFile || '';
  const templatePath = path.join( config.templateDirectory, templateFile );

  /**
   * Convert bytes to a human-readable file size string
   *
   * @param bytes - number of bytes
   * @returns formatted file size string (e.g. "1.5 MB")
   */
  function humanFileSize( bytes: number ): string {
    if ( bytes === 0 ) return '0 B';
    const k = 1024;
    const sizes = [ 'B', 'KB', 'MB', 'GB', 'TB' ];
    const i = Math.floor( Math.log( bytes ) / Math.log( k ) );
    return parseFloat( ( bytes / Math.pow( k, i ) ).toFixed( 2 ) ) + ' ' + sizes[ i ];
  }

  try {
    const stat = await fs.promises.stat( templatePath );
    res.json({
      name: templateFile,
      lastModified: stat.mtime.toLocaleString(),
      size: humanFileSize( stat.size )
    });
  } catch {
    res.json({ error: 'Template file not found' });
  }
});


/**
 * Run the folder creation script as a child process
 *
 * @param options - configuration for folder creation
 * @param options.month - month number (1-12)
 * @param options.year - full year (e.g. 2025)
 * @param options.write - whether to actually write files
 * @param options.foldersOnly - whether to create folders only (no template copy)
 * @returns result object with success status and message or error
 */
function runFolderCreation({ month, year, write, foldersOnly, overwrite }: { month: number, year: number, write: boolean, foldersOnly: boolean, overwrite?: boolean }): Promise<{ success: boolean, message?: string, error?: string }> {
  return new Promise( ( resolve ) => {
    let command: string;
    let args: string[];
    let cwd: string;
    let useShell = false;

    /**
     * Wrap a string in double quotes if it contains whitespace
     *
     * @param val - string to potentially quote
     * @returns quoted string if needed, otherwise the original
     */
    function quoteIfNeeded( val: string ): string {
      return /\s/.test( val ) ? `"${ val.replace( /"/g, '\"' ) }"` : val;
    }

    if ( process.env.NODE_ENV === 'development' ) {
      command = process.platform === 'win32' ? 'npx.cmd' : 'npx';
      args = [
        'ts-node',
        'src/index.ts',
        '--month', String( month ),
        '--year', String( year ),
        '--templateFile', quoteIfNeeded( String( config.templateFile || '' ) ),
        '--templateDirectory', quoteIfNeeded( String( config.templateDirectory || '' ) ),
        '--outputDirectory', quoteIfNeeded( String( config.outputDirectory || '' ) ),
        '--ext', quoteIfNeeded( String( config.ext || '' ) ),
        '--rootPath', quoteIfNeeded( String( config.rootPath || '' ) )
      ];
      if ( write ) args.push( '--write' );
      if ( foldersOnly !== false ) args.push( '--folders-only' );
      if ( overwrite ) args.push( '--overwrite' );
      cwd = path.resolve( __dirname, '..' ); // project root
      useShell = true; // npx on Windows sometimes needs shell
    } else {
      command = process.execPath; // node executable
      args = [
        'lib/index.js',
        '--month', String( month ),
        '--year', String( year ),
        '--templateFile', quoteIfNeeded( String( config.templateFile || '' ) ),
        '--templateDirectory', quoteIfNeeded( String( config.templateDirectory || '' ) ),
        '--outputDirectory', quoteIfNeeded( String( config.outputDirectory || '' ) ),
        '--ext', quoteIfNeeded( String( config.ext || '' ) ),
        '--rootPath', quoteIfNeeded( String( config.rootPath || '' ) )
      ];
      if ( write ) args.push( '--write' );
      if ( foldersOnly !== false ) args.push( '--folders-only' );
      if ( overwrite ) args.push( '--overwrite' );
      cwd = path.resolve( __dirname, '..' ); // project root
      useShell = false;
    }
    const child: child_process.ChildProcessWithoutNullStreams = child_process.spawn(
      command,
      args,
      {
        cwd,
        shell: useShell
      }
    );

    let output = '';
    let error = '';
    child.stdout.on( 'data', ( data: Buffer ) => { output += data.toString(); } );
    child.stderr.on( 'data', ( data: Buffer ) => { error += data.toString(); } );

    child.on( 'close', ( code: number ) => {
      if ( code === 0 ) {
        resolve({ success: true, message: output.trim() });
      } else {
        resolve({ success: false, error: ( error.trim() || output.trim() || 'Script failed.' ) });
      }
    });
  });
}


/**
 * API endpoint to run the folder creation script (folders only)
 */
app.post( '/api/run-folders', async ( req: Request, res: Response ) => {
  const { month, year, write } = req.body || {};
  if ( !month || !year ) {
    res.status( 400 ).json({ error: 'Month and year are required.' });
    return;
  }
  const result = await runFolderCreation({ month: Number( month ), year: Number( year ), write: !!write, foldersOnly: true });
  res.json( result );
});


/**
 * API endpoint to run the template copy script (full template copy)
 */
app.post( '/api/copy-template', async ( req: Request, res: Response ) => {
  const { month, year, write, overwrite } = req.body || {};
  if ( !month || !year ) {
    res.status( 400 ).json({ error: 'Month and year are required.' });
    return;
  }
  const result = await runFolderCreation({ month: Number( month ), year: Number( year ), write: !!write, foldersOnly: false, overwrite: !!overwrite });
  res.json( result );
});


/**
 * API endpoint to get notes file for a folder
 */
app.get( '/api/notes', ( req, res ) => {
  ( async () => {
    const folder = req.query.folder as string;
    if ( !folder ) return res.status( 400 ).send( 'Missing folder' );
    const folderPath = await resolveFolderPath( folder );
    if ( !folderPath ) return res.status( 404 ).send( 'Folder not found' );
    let files: string[] = [];
    try {
      files = await fs.promises.readdir( folderPath );
    } catch {
      return res.status( 404 ).send( 'Folder not found' );
    }
    const notesFile = files.find( f => /^\d{4}-?\d{2}-?\d{2}(\s+)?notes\.txt$/i.test( f ) );
    if ( !notesFile ) return res.status( 404 ).send( 'Notes file not found' );
    const notesPath = path.join( folderPath, notesFile );
    try {
      const text = await fs.promises.readFile( notesPath, 'utf8' );
      res.type( 'text/plain' ).send( text );
    } catch {
      res.status( 500 ).send( 'Could not read notes file' );
    }
  })();
});


/**
 * API endpoint to update notes file for a folder
 */
app.post( '/api/notes', function( req: any, res: any ) {
  const folder = req.query.folder as string;
  if ( !folder ) return res.status( 400 ).send( 'Missing folder' );
  ( async () => {
    const folderPath = await resolveFolderPath( folder );
    if ( !folderPath ) return res.status( 404 ).send( 'Folder not found' );
    const files = await fs.promises.readdir( folderPath );
    const notesFile = files.find( f => /^\d{4}-?\d{2}-?\d{2}(\s+)?notes\.txt$/i.test( f ) );
    if ( !notesFile ) { res.status( 404 ).send( 'Notes file not found' ); return; }
    const notesPath = path.join( folderPath, notesFile );
    let body = '';
    req.on( 'data', ( chunk: Buffer ) => { body += chunk.toString(); } );
    req.on( 'end', () => {
      fs.promises.writeFile( notesPath, body, 'utf8' )
        .then( () => res.send( 'OK' ) )
        .catch( () => res.status( 500 ).send( 'Could not save notes file' ) );
    });
  })().catch( () => res.status( 404 ).send( 'Folder not found' ) );
});


/**
 * API endpoint to get the YouTube URL for a Sunday folder.
 */
app.get( '/api/youtube-url', async ( req: Request, res: Response ) => {
  const folder = req.query.folder as string;
  if ( !folder || !/^\d{8}$/.test( folder ) ) {
    res.status( 400 ).json({ error: 'Missing or invalid folder (YYYYMMDD required)' });
    return;
  }

  try {
    const pool = getPool();
    const [ rows ] = await pool.query<RowDataPacket[]>(
      'SELECT url FROM youtube_urls WHERE sunday_date = ?',
      [ folder ]
    );
    res.json({ url: rows.length > 0 ? rows[ 0 ].url : null });
  } catch {
    res.json({ url: null });
  }
});


/**
 * API endpoint to archive a Sunday folder.
 *
 * Moves OUTPUT_DIRECTORY/YYYYMMDD or OUTPUT_DIRECTORY/YYYY Backlog/YYYYMMDD
 * into OUTPUT_DIRECTORY/YYYY/YYYYMMDD.
 * Returns 409 if the folder is already archived or not found.
 */
app.post( '/api/folders/:name/archive', async ( req: Request, res: Response ) => {
  const folder = req.params.name;
  if ( !folder || !/^\d{8}$/.test( folder ) ) {
    res.status( 400 ).json({ error: 'Missing or invalid folder (YYYYMMDD required)' });
    return;
  }

  const year = folder.slice( 0, 4 );
  const candidates = [
    path.join( config.outputDirectory, folder ),
    path.join( config.outputDirectory, `${ year } Backlog`, folder ),
  ];
  const archiveDir = path.join( config.outputDirectory, year, folder );

  let sourceDir: string | null = null;
  for ( const candidate of candidates ) {
    try {
      const stat = await fs.promises.stat( candidate );
      if ( stat.isDirectory() ) { sourceDir = candidate; break; }
    } catch {}
  }

  if ( !sourceDir ) {
    res.status( 409 ).json({ error: 'Folder is already archived or does not exist' });
    return;
  }

  try {
    await fs.promises.mkdir( path.join( config.outputDirectory, year ), { recursive: true } );
    await fs.promises.rename( sourceDir, archiveDir );
    console.log( `[ARCHIVE] Moved ${ folder } to ${ year }/` );
    io.emit( 'folder-changes', { name: folder } );
    res.json({ ok: true });
  } catch ( err ) {
    console.error( '[ARCHIVE] Failed to move folder:', err );
    res.status( 500 ).json({ error: 'Failed to archive folder' });
  }
});


/**
 * API endpoint to save/update the YouTube URL for a Sunday folder.
 */
app.put( '/api/youtube-url', async ( req: Request, res: Response ) => {
  const { folder, url } = req.body || {};
  if ( !folder || !/^\d{8}$/.test( folder ) ) {
    res.status( 400 ).json({ error: 'Missing or invalid folder (YYYYMMDD required)' });
    return;
  }

  try {
    const pool = getPool();
    if ( !url || url.trim() === '' ) {
      await pool.query( 'DELETE FROM youtube_urls WHERE sunday_date = ?', [ folder ] );
    } else {
      await pool.query(
        'INSERT INTO youtube_urls (sunday_date, url) VALUES (?, ?) ON DUPLICATE KEY UPDATE url = VALUES(url)',
        [ folder, url.trim() ]
      );
    }
    // Move folder from Backlog to archive when YouTube URL is set
    if ( url && url.trim() !== '' ) {
      const year = folder.slice( 0, 4 );
      const backlogDir = path.join( config.outputDirectory, `${ year } Backlog`, folder );
      const archiveDir = path.join( config.outputDirectory, year, folder );
      try {
        const stat = await fs.promises.stat( backlogDir );
        if ( stat.isDirectory() ) {
          await fs.promises.mkdir( path.join( config.outputDirectory, year ), { recursive: true } );
          await fs.promises.rename( backlogDir, archiveDir );
          console.log( `[YOUTUBE] Moved ${ folder } from Backlog to ${ year }/` );
        }
      } catch {
        // Folder not in Backlog — nothing to move
      }
    }

    io.emit( 'folder-changes', { name: folder } );
    res.json({ ok: true });
  } catch ( err ) {
    res.status( 500 ).json({ error: 'Failed to save YouTube URL' });
  }
});


/**
 * API endpoint to list sermon-eligible PPTX files in a Sunday folder.
 * Returns all files matching the configured extension except the template copy (YYYYMMDD.ext).
 */
app.get( '/api/sermon-pptx-files', async ( req: Request, res: Response ) => {
  const folder = req.query.folder as string;
  if ( !folder || !/^\d{8}$/.test( folder ) ) {
    res.status( 400 ).json({ error: 'Missing or invalid folder (YYYYMMDD required)' });
    return;
  }

  const folderPath = await resolveFolderPath( folder );
  if ( !folderPath ) {
    res.json({ files: [] });
    return;
  }

  try {
    const files = await fs.promises.readdir( folderPath );
    const ext = '.' + config.ext.toLowerCase();
    const templateName = ( folder + ext ).toLowerCase();
    const pptxFiles = files.filter( f => {
      const lower = f.toLowerCase();
      return lower.endsWith( ext ) && lower !== templateName;
    });
    res.json({ files: pptxFiles });
  } catch {
    res.json({ files: [] });
  }
});


/**
 * API endpoint to get the sermon info for a Sunday folder.
 */
app.get( '/api/sermon-info', async ( req: Request, res: Response ) => {
  const folder = req.query.folder as string;
  if ( !folder || !/^\d{8}$/.test( folder ) ) {
    res.status( 400 ).json({ error: 'Missing or invalid folder (YYYYMMDD required)' });
    return;
  }

  try {
    const pool = getPool();
    const [ rows ] = await pool.query<RowDataPacket[]>(
      'SELECT pptx_file, title, speaker FROM sermon_info WHERE sunday_date = ?',
      [ folder ]
    );
    res.json({
      pptxFile: rows.length > 0 ? rows[ 0 ].pptx_file : null,
      title: rows.length > 0 ? rows[ 0 ].title : null,
      speaker: rows.length > 0 ? rows[ 0 ].speaker : null,
    });
  } catch {
    res.json({ pptxFile: null, title: null, speaker: null });
  }
});


/**
 * API endpoint to save or update sermon info for a Sunday folder.
 */
app.put( '/api/sermon-info', async ( req: Request, res: Response ) => {
  const { folder, pptxFile, title, speaker } = req.body || {};
  if ( !folder || !/^\d{8}$/.test( folder ) ) {
    res.status( 400 ).json({ error: 'Missing or invalid folder (YYYYMMDD required)' });
    return;
  }

  try {
    const pool = getPool();
    const trimmedTitle = title ? title.trim() : null;
    const trimmedSpeaker = speaker ? speaker.trim() : null;
    const trimmedFile = pptxFile ? pptxFile.trim() : null;

    if ( !trimmedFile && !trimmedTitle && !trimmedSpeaker ) {
      await pool.query( 'DELETE FROM sermon_info WHERE sunday_date = ?', [ folder ] );
    } else {
      await pool.query(
        `INSERT INTO sermon_info (sunday_date, pptx_file, title, speaker)
         VALUES (?, ?, ?, ?)
         ON DUPLICATE KEY UPDATE pptx_file = VALUES(pptx_file), title = VALUES(title), speaker = VALUES(speaker)`,
        [ folder, trimmedFile, trimmedTitle, trimmedSpeaker ]
      );
    }
    io.emit( 'folder-changes', { name: folder } );
    res.json({ ok: true });
  } catch ( err ) {
    console.error( '[SERMON INFO] Error:', err );
    res.status( 500 ).json({ error: 'Failed to save sermon info' });
  }
});


/**
 * API endpoint to list all stored speakers.
 */
app.get( '/api/speakers', async ( _req: Request, res: Response ) => {
  try {
    const pool = getPool();
    const [ rows ] = await pool.query<RowDataPacket[]>( 'SELECT id, name FROM speakers ORDER BY name' );
    res.json( rows.map( r => ({ id: r.id, name: r.name }) ) );
  } catch {
    res.json([]);
  }
});


/**
 * API endpoint to add a new speaker to the reusable list.
 */
app.post( '/api/speakers', async ( req: Request, res: Response ) => {
  const { name } = req.body || {};
  if ( !name || !name.trim() ) {
    res.status( 400 ).json({ error: 'Missing speaker name' });
    return;
  }

  try {
    const pool = getPool();
    const [ result ] = await pool.query<ResultSetHeader>(
      'INSERT IGNORE INTO speakers (name) VALUES (?)',
      [ name.trim() ]
    );
    res.json({ ok: true, id: result.insertId || null });
  } catch ( err ) {
    console.error( '[SPEAKERS] Error:', err );
    res.status( 500 ).json({ error: 'Failed to add speaker' });
  }
});


/**
 * API endpoint to approve a presentation by renaming the TODO .lnk shortcut in the template directory.
 */
app.put( '/api/folders/:name/approve', async ( req: Request, res: Response ) => {
  const folderName = req.params.name;
  if ( !folderName || !/^\d{8}$/.test( folderName ) ) {
    res.status( 400 ).json({ error: 'Invalid folder name (YYYYMMDD required)' });
    return;
  }

  const dateStr = folderName.replace( /(\d{4})(\d{2})(\d{2})/, '$1-$2-$3' );
  const todoLnk = `${ dateStr } TODO.lnk`;
  const doneLnk = `${ dateStr }.lnk`;
  const oldPath = path.join( config.templateDirectory, todoLnk );
  const newPath = path.join( config.templateDirectory, doneLnk );

  try {
    await fs.promises.access( oldPath );
  } catch {
    // No TODO shortcut — already approved or missing
    res.json({ ok: true, message: 'Already approved' });
    return;
  }

  try {
    await fs.promises.rename( oldPath, newPath );
    console.log( `[APPROVE] Renamed "${ todoLnk }" → "${ doneLnk }" in template directory` );
    io.emit( 'folder-changes', { name: folderName } );
    res.json({ ok: true, oldName: todoLnk, newName: doneLnk });
  } catch ( err ) {
    console.error( `[APPROVE] Failed to rename shortcut for ${ folderName }:`, err );
    res.status( 500 ).json({ error: 'Failed to rename shortcut' });
  }
});


/**
 * Cancel approval of a presentation by adding "TODO" back to the shortcut filename.
 * Renames "YYYY-MM-DD.lnk" → "YYYY-MM-DD TODO.lnk" in the template directory.
 */
app.put( '/api/folders/:name/unapprove', async ( req: Request, res: Response ) => {
  const folderName = req.params.name;
  if ( !folderName || !/^\d{8}$/.test( folderName ) ) {
    res.status( 400 ).json({ error: 'Invalid folder name (YYYYMMDD required)' });
    return;
  }

  const dateStr = folderName.replace( /(\d{4})(\d{2})(\d{2})/, '$1-$2-$3' );
  const doneLnk = `${ dateStr }.lnk`;
  const todoLnk = `${ dateStr } TODO.lnk`;
  const oldPath = path.join( config.templateDirectory, doneLnk );
  const newPath = path.join( config.templateDirectory, todoLnk );

  try {
    await fs.promises.access( oldPath );
  } catch {
    // No approved shortcut — already unapproved or missing
    res.json({ ok: true, message: 'Already unapproved' });
    return;
  }

  try {
    await fs.promises.rename( oldPath, newPath );
    console.log( `[UNAPPROVE] Renamed "${ doneLnk }" → "${ todoLnk }" in template directory` );
    io.emit( 'folder-changes', { name: folderName } );
    res.json({ ok: true, oldName: doneLnk, newName: todoLnk });
  } catch ( err ) {
    console.error( `[UNAPPROVE] Failed to rename shortcut for ${ folderName }:`, err );
    res.status( 500 ).json({ error: 'Failed to rename shortcut' });
  }
});


/**
 * List video files in a Sunday folder's Vids/ subdirectory.
 * Returns MKV and MP4 files (excluding production MP4s) with file size and type.
 */
app.get( '/api/folders/:name/video-files', async ( req: Request, res: Response ) => {
  const folderName = req.params.name;
  if ( !folderName || !/^\d{8}$/.test( folderName ) ) {
    res.status( 400 ).json({ error: 'Invalid folder name (YYYYMMDD required)' });
    return;
  }

  const folderPath = await resolveFolderPath( folderName );
  if ( !folderPath ) {
    res.json( [] );
    return;
  }
  const vidsDir = path.join( folderPath, 'Vids' );
  try {
    const files = await fs.promises.readdir( vidsDir );
    const videoFiles = [];

    for ( const f of files ) {
      const lower = f.toLowerCase();
      const isVideo = lower.endsWith( '.mp4' ) || lower.endsWith( '.mkv' );
      const isProduction = /^\d{8}-production\.mp4$/i.test( f );
      if ( !isVideo || isProduction ) continue;

      const stat = await fs.promises.stat( path.join( vidsDir, f ) );
      videoFiles.push({
        filename: f,
        size: stat.size,
        type: lower.endsWith( '.mkv' ) ? 'mkv' : 'mp4',
      });
    }

    res.json( videoFiles );
  } catch {
    res.json( [] );
  }
});


/**
 * Generate a Kdenlive project file from selected video files.
 * Expects cameraFile (MP4 with 1 audio + 1 video) and optional obsFile (MKV with 3 audio + 1 video).
 */
app.post( '/api/folders/:name/create-kdenlive', async ( req: Request, res: Response ) => {
  const folderName = req.params.name;
  if ( !folderName || !/^\d{8}$/.test( folderName ) ) {
    res.status( 400 ).json({ error: 'Invalid folder name (YYYYMMDD required)' });
    return;
  }

  const { obsFile, cameraFile } = req.body;
  if ( !cameraFile ) {
    res.status( 400 ).json({ error: 'cameraFile is required' });
    return;
  }

  const folderPath = await resolveFolderPath( folderName );
  if ( !folderPath ) {
    res.status( 404 ).json({ error: 'Folder not found' });
    return;
  }
  const vidsDir = path.join( folderPath, 'Vids' );

  // Validate files exist
  try {
    if ( obsFile ) {
      await fs.promises.access( path.join( vidsDir, obsFile ) );
    }
    await fs.promises.access( path.join( vidsDir, cameraFile ) );
  } catch {
    res.status( 404 ).json({ error: 'Video file(s) not found in Vids/' });
    return;
  }

  const outputFile = path.join( vidsDir, `${ folderName }.kdenlive` );

  // Don't overwrite existing project
  try {
    await fs.promises.access( outputFile );
    res.status( 409 ).json({ error: 'Kdenlive project already exists' });
    return;
  } catch {
    // File doesn't exist — good, continue
  }

  try {
    // Query videos table for accurate duration
    let duration: string | undefined;
    try {
      const filesToCheck = [ cameraFile, obsFile ].filter( Boolean );
      const relPaths = filesToCheck.map( f => toJobPath( path.join( vidsDir, f! ) ) );
      const videos = await getVideosByPaths( relPaths );
      let maxSecs = 0;
      for ( const v of videos ) {
        if ( v.duration_secs && v.duration_secs > maxSecs ) {
          maxSecs = v.duration_secs;
          duration = v.duration_timecode ?? undefined;
        }
      }
    } catch {
      // ffprobe metadata not available — use default duration
    }

    const { generateKdenlive } = require( './kdenlive-generator' );
    const xml = generateKdenlive({
      sundayDate: folderName,
      rootPath: vidsDir.replace( /\\/g, '/' ),
      obsFile: obsFile || undefined,
      cameraFile,
      duration,
    });

    await fs.promises.writeFile( outputFile, xml, 'utf-8' );
    console.log( `[KDENLIVE] Generated project: ${ outputFile }` );
    io.emit( 'folder-changes', { name: folderName } );
    res.json({ ok: true, path: `Vids/${ folderName }.kdenlive` });
  } catch ( err ) {
    console.error( `[KDENLIVE] Failed to generate project for ${ folderName }:`, err );
    res.status( 500 ).json({ error: 'Failed to generate Kdenlive project' });
  }
});


/**
 * API endpoint to approve a Kdenlive project and queue a transcode job.
 * Finds the .kdenlive file in Vids/, creates a transcode job that will
 * render it to YYYYMMDD-production.mp4, which then auto-chains to transcription.
 */
app.post( '/api/folders/:name/approve-kdenlive', async ( req: Request, res: Response ) => {
  const folderName = req.params.name;
  if ( !folderName || !/^\d{8}$/.test( folderName ) ) {
    res.status( 400 ).json({ error: 'Invalid folder name (YYYYMMDD required)' });
    return;
  }

  const folderPath = await resolveFolderPath( folderName );
  if ( !folderPath ) {
    res.status( 404 ).json({ error: 'Folder not found' });
    return;
  }
  const vidsDir = path.join( folderPath, 'Vids' );

  try {
    // Find the .kdenlive file
    const vidsFiles = await fs.promises.readdir( vidsDir );
    const kdenliveFile = vidsFiles.find( f => f.toLowerCase().endsWith( '.kdenlive' ) );

    if ( !kdenliveFile ) {
      res.status( 404 ).json({ error: 'No Kdenlive project found in Vids/' });
      return;
    }

    const kdenlivePath = toJobPath( path.join( vidsDir, kdenliveFile ) );
    const productionPath = toJobPath( path.join( vidsDir, `${ folderName }-production.mp4` ) );

    // Check if production MP4 already exists
    try {
      await fs.promises.access( path.join( vidsDir, `${ folderName }-production.mp4` ) );
      res.status( 409 ).json({ error: 'Production MP4 already exists' });
      return;
    } catch {
      // Good — production doesn't exist yet
    }

    const job = await createJobIfNotExists({
      type: 'transcode',
      sunday_date: folderName,
      input_path: kdenlivePath,
      output_path: productionPath,
      metadata: { kdenlive_file: kdenliveFile },
    });

    if ( !job ) {
      res.status( 409 ).json({ error: 'A transcode job already exists for this date' });
      return;
    }

    io.emit( 'job:created', job );
    console.log( `[KDENLIVE] Queued transcode job #${ job.id } for ${ folderName }` );

    // Create a held ffprobe job for the production file (released when transcode completes)
    const heldProbe = await createHeldProbeJob( folderName, productionPath, job.id );
    if ( heldProbe ) {
      io.emit( 'job:created', heldProbe );
      console.log( `[KDENLIVE] Created held ffprobe job #${ heldProbe.id } for ${ folderName }` );
    }

    res.json({ ok: true, jobId: job.id });
  } catch ( err ) {
    console.error( `[KDENLIVE] Failed to approve Kdenlive for ${ folderName }:`, err );
    res.status( 500 ).json({ error: 'Failed to queue transcode job' });
  }
});


/**
 * API endpoint to get the bible verse for a given month.
 */
app.get( '/api/verse', async ( req: Request, res: Response ) => {
  const year = parseInt( req.query.year as string );
  const month = parseInt( req.query.month as string );
  if ( !year || !month ) {
    res.status( 400 ).json({ error: 'Missing year or month' });
    return;
  }

  try {
    const pool = getPool();
    const [ rows ] = await pool.query<RowDataPacket[]>(
      'SELECT reference, text FROM bible_verses WHERE year = ? AND month = ?',
      [ year, month ]
    );
    res.json({
      reference: rows.length > 0 ? rows[ 0 ].reference : '',
      text: rows.length > 0 ? ( rows[ 0 ].text || '' ) : '',
    });
  } catch ( err ) {
    res.json({ reference: '', text: '' });
  }
});


/**
 * API endpoint to save the bible verse for a given month.
 */
app.post( '/api/verse', async ( req: Request, res: Response ) => {
  const { year, month, reference, text } = req.body || {};
  if ( !year || !month ) {
    res.status( 400 ).json({ error: 'Missing year or month' });
    return;
  }

  try {
    const pool = getPool();
    if ( ( !reference || reference.trim() === '' ) && ( !text || text.trim() === '' ) ) {
      await pool.query( 'DELETE FROM bible_verses WHERE year = ? AND month = ?', [ year, month ] );
    } else {
      await pool.query(
        'INSERT INTO bible_verses (year, month, reference, text) VALUES (?, ?, ?, ?) ON DUPLICATE KEY UPDATE reference = VALUES(reference), text = VALUES(text)',
        [ year, month, ( reference || '' ).trim(), text ? text.trim() : null ]
      );
    }
    res.json({ ok: true });
  } catch ( err ) {
    res.status( 500 ).json({ error: 'Failed to save verse' });
  }
});


/**
 * API endpoint to get the closing song for a given month.
 */
app.get( '/api/closing-song', async ( req: Request, res: Response ) => {
  const year = parseInt( req.query.year as string );
  const month = parseInt( req.query.month as string );
  if ( !year || !month ) {
    res.status( 400 ).json({ error: 'Missing year or month' });
    return;
  }

  try {
    const song = await getClosingSong( year, month );
    res.json({ song });
  } catch ( err ) {
    res.json({ song: null });
  }
});


/**
 * API endpoint to set the closing song for a given month.
 * Creates a "closing song" shortcut in every Sunday folder for that month.
 */
app.post( '/api/closing-song', ( req: Request, res: Response ) => {
  ( async () => {
    try {
      const { year, month, song } = req.body || {};
      if ( !year || !month || !song ) {
        return res.status( 400 ).json({ error: 'Missing year, month, or song' });
      }

      // Persist the selection in song_selections
      await setClosingSong( year, month, song );

      // Create shortcuts in every Sunday folder for this month
      const sundays = sundaysInMonth( month, year );

      for ( const day of sundays ) {
        const mm = String( month ).padStart( 2, '0' );
        const dd = String( day ).padStart( 2, '0' );
        const dateStr = `${ year }${ mm }${ dd }`;
        const weekFolder = await resolveFolderPath( dateStr );

        if ( !weekFolder ) continue;

        // Remove existing closing song shortcuts
        deleteShortcutsByPrefix( weekFolder, 'closing song ' );

        // Find the song file in the library
        const songFile = findSongFile( song, config.songsDirectory );
        if ( !songFile ) {
          console.log( `[CLOSING SONG] Song file not found in library for ${ song.number || '' } ${ song.name }` );
          continue;
        }

        const shortcutName = `closing song ${ song.number ? song.number + ' - ' : '' }${ song.name }`.replace( /[\\/:*?"<>|]/g, '_' ) + ' - Shortcut.lnk';
        const shortcutPath = path.join( weekFolder, shortcutName );
        const desc = `Closing Song: ${ song.number ? song.number + ' - ' : '' }${ song.name }`;
        const rootPath = config.rootPath;

        const created = await createShortcut( shortcutPath, desc, songFile, rootPath );
        if ( created ) {
          console.log( `[CLOSING SONG] Created shortcut in ${ dateStr }` );
        } else {
          console.error( `[CLOSING SONG] Failed to create shortcut in ${ dateStr }` );
        }
      }

      res.json({ ok: true });
    } catch ( err ) {
      console.error( '[CLOSING SONG] Error:', err );
      res.status( 500 ).json({ error: 'Failed to save closing song' });
    }
  })();
});


/**
 * API endpoint to clear the closing song for a given month.
 * Removes the song_selections record and deletes closing song shortcuts from all Sunday folders.
 */
app.delete( '/api/closing-song', ( req: Request, res: Response ) => {
  ( async () => {
    try {
      const year = parseInt( req.query.year as string );
      const month = parseInt( req.query.month as string );
      if ( !year || !month ) {
        return res.status( 400 ).json({ error: 'Missing year or month' });
      }

      // Remove the selection from song_selections
      await setClosingSong( year, month, null );

      // Remove closing song shortcuts from all Sunday folders
      const sundays = sundaysInMonth( month, year );

      for ( const day of sundays ) {
        const mm = String( month ).padStart( 2, '0' );
        const dd = String( day ).padStart( 2, '0' );
        const dateStr = `${ year }${ mm }${ dd }`;
        const weekFolder = await resolveFolderPath( dateStr );

        if ( weekFolder ) {
          deleteShortcutsByPrefix( weekFolder, 'closing song ' );
        }
      }

      res.json({ ok: true });
    } catch ( err ) {
      console.error( '[CLOSING SONG] Error clearing:', err );
      res.status( 500 ).json({ error: 'Failed to clear closing song' });
    }
  })();
});


/**
 * API endpoint to list songs in songsDirectory, enriched with DB stats.
 */
app.get( '/api/songs', async ( req: Request, res: Response ) => {
  const songsDir = config.songsDirectory;
  let songFiles: {
    id?: number;
    name: string;
    number?: string;
    book?: string;
    bookIcon?: string | null;
    filePath?: string;
    lastModified?: string;
    ccli: string | null;
    license: string | null;
    lastUsed?: string;
    totalUsed?: number;
  }[] = [];
  const ext = '.' + ( config.ext || 'pptx' ).toLowerCase();

  /**
   * Recursively walk a directory and collect song files
   *
   * @param dir - absolute path to the directory to walk
   * @param relBase - relative base path for tracking book/source
   */
  async function walk( dir: string, relBase: string = '' ) {
    let files: string[] = [];
    try {
      files = await fs.promises.readdir( dir );
    } catch ( err ) {
      return;
    }
    for ( const f of files ) {
      const fullPath = path.join( dir, f );
      const relPath = relBase ? path.join( relBase, f ) : f;
      let stat;
      try {
        stat = await fs.promises.stat( fullPath );
      } catch { continue; }
      if ( stat.isDirectory() ) {
        if ( f.toLowerCase() === 'archive' || f.toLowerCase() === 'icons' ) continue;
        await walk( fullPath, relPath );
      } else if ( f.toLowerCase().endsWith( ext ) ) {
        // Extract song number if filename starts with exactly 3 digits, optionally followed by space or dash
        // Example: 123 - Song Name.pptx or 123- Song Name.pptx or 123-Name.pptx
        const match = f.match( /^(\d{3})\s*-?\s*/ );
        const number = match ? match[ 1 ] : undefined;
        // Book/source is the top-level folder (first part of relPath)
        const book = relPath.split( path.sep )[ 0 ];
        // Remove number and extension from name
        let name = f.replace( /^(\d{3})\s*-?\s*/, '' ).replace( new RegExp( ext + '$', 'i' ), '' );
        // Format last modified date
        const lastModified = stat.mtime.toLocaleString();
        // Build path with env variable prefix for easy copy/paste into Run
        const resolvedRoot = resolveToAbsolutePath( config.rootPath || '%OneDriveConsumer%' ).replace( /\/$/, '' );
        const normalFull = fullPath.replace( /\//g, '\\' );
        const normalRoot = resolvedRoot.replace( /\//g, '\\' );
        const varPath = normalFull.toLowerCase().startsWith( normalRoot.toLowerCase() )
          ? ( config.rootPath || '%OneDriveConsumer%' ) + normalFull.slice( normalRoot.length )
          : normalFull;
        songFiles.push({ name, number, book, filePath: varPath, lastModified, ccli: null, license: null });
      }
    }
  }

  await walk( songsDir );

  // Enrich with CCLI, last used date, and total usage from the database
  try {
    const pool = getPool();
    const [ rows ] = await pool.query<RowDataPacket[]>(
      `SELECT s.id, s.normalized_name, s.number, s.book, s.ccli, s.license,
              COUNT( sel.id ) AS total_used,
              MAX( sel.sunday_date ) AS last_used
       FROM songs s
       LEFT JOIN song_selections sel ON sel.song_id = s.id AND sel.removed_at IS NULL
       GROUP BY s.id`
    );

    // Build lookup map: "normalized|number|book" → { id, ccli, lastUsed, totalUsed }
    const statsMap = new Map<string, { id: number; ccli?: string; license?: string; lastUsed?: string; totalUsed: number }>();
    for ( const row of rows ) {
      const key = `${ row.normalized_name }|${ row.number || '' }|${ ( row.book || '' ).toLowerCase() }`;
      statsMap.set( key, {
        id: row.id,
        ccli: row.ccli || undefined,
        license: row.license || undefined,
        lastUsed: row.last_used || undefined,
        totalUsed: Number( row.total_used ) || 0,
      });
    }

    // Merge DB stats into filesystem song list, inserting missing songs
    const missing: typeof songFiles = [];
    for ( const song of songFiles ) {
      const normalized = normalizeSongName( song.name );
      const key = `${ normalized }|${ song.number || '' }|${ ( song.book || '' ).toLowerCase() }`;
      const stats = statsMap.get( key );
      if ( stats ) {
        song.id = stats.id;
        song.ccli = stats.ccli || null;
        song.license = stats.license || null;
        song.lastUsed = stats.lastUsed;
        song.totalUsed = stats.totalUsed;
      } else {
        song.totalUsed = 0;
        missing.push( song );
      }
    }

    // Seed missing songs into the database
    for ( const song of missing ) {
      try {
        const id = await findOrCreateSong( song.name, song.number, song.book, song.filePath );
        song.id = id;
      } catch {
        // Non-fatal — song just won't have a DB record yet
      }
    }
  } catch {
    // DB not available — return songs without stats
  }

  // Deduplicate by normalized name + number + book (e.g. "It's" vs "Its" on disk)
  const seen = new Map<string, number>();
  songFiles = songFiles.filter( ( song, idx ) => {
    const normalized = normalizeSongName( song.name );
    const key = `${ normalized }|${ song.number || '' }|${ ( song.book || '' ).toLowerCase() }`;
    if ( seen.has( key ) ) {
      // Keep the one with a DB id; if both have one, keep the first
      const prevIdx = seen.get( key )!;
      if ( !songFiles[ prevIdx ].id && song.id ) {
        songFiles[ prevIdx ] = song;
      }
      return false;
    }
    seen.set( key, idx );
    return true;
  });

  // Attach book icons from the books table
  try {
    const pool = getPool();
    const [ bookRows ] = await pool.query<RowDataPacket[]>(
      'SELECT name, icon FROM books WHERE enabled = TRUE AND icon IS NOT NULL'
    );
    const iconsDir = path.join( config.songsDirectory, 'icons' );
    const iconMap = new Map<string, string>();
    for ( const row of bookRows ) {
      // Only include if the icon file actually exists on disk
      try {
        fs.accessSync( path.join( iconsDir, row.icon ) );
        iconMap.set( row.name.toLowerCase(), row.icon );
      } catch {
        // File doesn't exist — skip
      }
    }
    for ( const song of songFiles ) {
      const icon = iconMap.get( ( song.book || '' ).toLowerCase() );
      if ( icon ) {
        song.bookIcon = icon;
      }
    }
  } catch {
    // DB not available — songs returned without book icons
  }

  // Sort by book, then song number, then name
  songFiles.sort( ( a, b ) => {
    const bookCmp = ( a.book || '' ).localeCompare( b.book || '' );
    if ( bookCmp !== 0 ) return bookCmp;
    const numCmp = ( a.number || '' ).localeCompare( b.number || '', undefined, { numeric: true } );
    if ( numCmp !== 0 ) return numCmp;
    return a.name.localeCompare( b.name );
  });

  res.json( songFiles );
});


/**
 * Build a song history response from a song ID.
 *
 * @param songId - the song primary key
 * @returns JSON-ready history object
 */
async function buildSongHistory( songId: number ) {
  const pool = getPool();

  const [ songRows ] = await pool.query<RowDataPacket[]>(
    `SELECT s.id, s.name, s.number, s.book, s.ccli, s.license, b.url_template
     FROM songs s
     LEFT JOIN books b ON b.name = s.book AND b.enabled = TRUE
     WHERE s.id = ?`,
    [ songId ]
  );
  if ( songRows.length === 0 ) return null;
  const song = songRows[ 0 ];

  const [ histRows ] = await pool.query<RowDataPacket[]>(
    `SELECT sunday_date, slot_type, slot_number, slot_suffix
     FROM song_selections
     WHERE song_id = ? AND removed_at IS NULL
     ORDER BY sunday_date DESC, slot_type, slot_number`,
    [ songId ]
  );

  const history = histRows.map( ( r: any ) => ({
    sundayDate: r.sunday_date,
    slotType: r.slot_type,
    slotNumber: r.slot_number,
    slotSuffix: r.slot_suffix || null,
  }));

  return {
    id: song.id,
    name: song.name,
    number: song.number || null,
    book: song.book || null,
    ccli: song.ccli || null,
    license: song.license || null,
    urlTemplate: song.url_template || null,
    totalUsed: history.length,
    firstUsed: history.length > 0 ? history[ history.length - 1 ].sundayDate : null,
    lastUsed: history.length > 0 ? history[ 0 ].sundayDate : null,
    history,
  };
}


/**
 * Get song usage history by ID.
 *
 * @param id - song primary key
 */
app.get( '/api/songs/:id/history', async ( req: Request, res: Response ) => {
  const songId = parseInt( req.params.id );
  if ( isNaN( songId ) || songId <= 0 ) {
    res.status( 400 ).json({ error: 'Invalid song ID' });
    return;
  }

  try {
    const result = await buildSongHistory( songId );
    if ( !result ) {
      res.status( 404 ).json({ error: 'Song not found' });
      return;
    }
    res.json( result );
  } catch {
    res.status( 503 ).json({ error: 'Database not available' });
  }
});


/**
 * Get song usage history by name (legacy fallback).
 *
 * @query name - song display name (required)
 * @query number - song number (optional)
 * @query book - book/source folder name (optional)
 */
app.get( '/api/songs/history', async ( req: Request, res: Response ) => {
  const name = req.query.name as string;
  if ( !name ) {
    res.status( 400 ).json({ error: 'name parameter is required' });
    return;
  }

  const number = req.query.number as string | undefined;
  const book = req.query.book as string | undefined;
  const normalized = normalizeSongName( name );

  try {
    const pool = getPool();

    // Tiered song lookup (read-only, mirrors findOrCreateSong tiers)
    let songId: number | null = null;

    if ( number && book ) {
      const [ rows ] = await pool.query<RowDataPacket[]>(
        'SELECT id FROM songs WHERE number = ? AND normalized_name = ? AND book = ?',
        [ number, normalized, book ]
      );
      if ( rows.length === 1 ) songId = rows[ 0 ].id;
    }

    if ( songId === null && number ) {
      const [ rows ] = await pool.query<RowDataPacket[]>(
        'SELECT id FROM songs WHERE number = ? AND normalized_name = ?',
        [ number, normalized ]
      );
      if ( rows.length === 1 ) songId = rows[ 0 ].id;
    }

    if ( songId === null && book ) {
      const [ rows ] = await pool.query<RowDataPacket[]>(
        'SELECT id FROM songs WHERE normalized_name = ? AND book = ?',
        [ normalized, book ]
      );
      if ( rows.length === 1 ) songId = rows[ 0 ].id;
    }

    if ( songId === null ) {
      const [ rows ] = await pool.query<RowDataPacket[]>(
        'SELECT id FROM songs WHERE normalized_name = ?',
        [ normalized ]
      );
      if ( rows.length === 1 ) songId = rows[ 0 ].id;
    }

    if ( songId === null ) {
      res.json({
        id: null,
        name,
        number: number || null,
        book: book || null,
        ccli: null,
        license: null,
        totalUsed: 0,
        firstUsed: null,
        lastUsed: null,
        history: [],
      });
      return;
    }

    const result = await buildSongHistory( songId );
    res.json( result );
  } catch {
    res.status( 503 ).json({ error: 'Database not available' });
  }
});


/**
 * Update CCLI and/or license on a song.
 * If id is 0, finds or creates the song record from name/number/book in the body.
 *
 * @param id - song primary key (or 0 to find-or-create)
 * @body ccli - CCLI number (string or null)
 * @body license - license info (string or null)
 * @body name - song display name (required when id is 0)
 * @body number - optional song number (used for find-or-create)
 * @body book - optional book name (used for find-or-create)
 */
app.put( '/api/songs/:id', async ( req: Request, res: Response ) => {
  const { ccli, license, name, number, book, filePath } = req.body || {};

  try {
    let songId = parseInt( req.params.id );

    // If id is 0 or invalid, find-or-create the song from name/number/book
    if ( !songId || isNaN( songId ) ) {
      if ( !name ) {
        res.status( 400 ).json({ error: 'name is required when song has no id' });
        return;
      }
      songId = await findOrCreateSong( name, number || undefined, book || undefined, filePath || undefined );
    }

    const pool = getPool();
    await pool.query(
      'UPDATE songs SET ccli = ?, license = ? WHERE id = ?',
      [ ccli || null, license || null, songId ]
    );
    res.json({ ok: true, id: songId });
  } catch {
    res.status( 503 ).json({ error: 'Database not available' });
  }
});


/**
 * Get all enabled books from the database.
 */
app.get( '/api/books', async ( _req: Request, res: Response ) => {
  try {
    const pool = getPool();
    const [ rows ] = await pool.query<RowDataPacket[]>(
      'SELECT id, name, icon FROM books WHERE enabled = TRUE ORDER BY name'
    );

    // Only include icon if the file actually exists on disk
    const iconsDir = path.join( config.songsDirectory, 'icons' );
    const booksWithIcons = await Promise.all( rows.map( async ( row ) => {
      if ( row.icon ) {
        try {
          await fs.promises.access( path.join( iconsDir, row.icon ) );
          return row;
        } catch {
          return { ...row, icon: null };
        }
      }
      return row;
    }) );

    res.json( booksWithIcons );
  } catch {
    // DB not available — fall back to scanning the songs directory
    try {
      const entries = await fs.promises.readdir( config.songsDirectory, { withFileTypes: true } );
      const books = entries
        .filter( e => e.isDirectory() )
        .map( ( e, i ) => ({ id: i + 1, name: e.name }) )
        .sort( ( a, b ) => a.name.localeCompare( b.name ) );
      res.json( books );
    } catch {
      res.json( [] );
    }
  }
});


/**
 * Create a new book (directory + DB row).
 */
app.post( '/api/books', async ( req: Request, res: Response ) => {
  const { name } = req.body || {};
  if ( !name || typeof name !== 'string' || !name.trim() ) {
    res.status( 400 ).json({ error: 'Book name is required' });
    return;
  }

  const trimmed = name.trim();

  try {
    // Create the folder on disk
    await fs.promises.mkdir( path.join( config.songsDirectory, trimmed ), { recursive: true } );

    // Insert into DB
    const pool = getPool();
    const [ result ] = await pool.query<ResultSetHeader>(
      'INSERT INTO books (name) VALUES (?) ON DUPLICATE KEY UPDATE enabled = TRUE',
      [ trimmed ]
    );
    const id = result.insertId || 0;
    res.json({ id, name: trimmed, enabled: true });
  } catch ( err: any ) {
    console.error( '[BOOKS] Error creating book:', err );
    res.status( 500 ).json({ error: 'Failed to create book' });
  }
});


/**
 * Create a new song file by copying the song template to the target book folder.
 */
app.post( '/api/songs/create', async ( req: Request, res: Response ) => {
  const { book, name, number, ccli } = req.body || {};

  if ( !name || typeof name !== 'string' || !name.trim() ) {
    res.status( 400 ).json({ error: 'Song name is required' });
    return;
  }
  if ( !book || typeof book !== 'string' || !book.trim() ) {
    res.status( 400 ).json({ error: 'Book is required' });
    return;
  }

  const songName = name.trim();
  const bookName = book.trim();
  const songNumber = number ? String( number ).trim() : '';
  const ext = '.' + ( config.ext || 'pptx' );

  // Build filename
  const filename = songNumber
    ? `${ songNumber } - ${ songName }${ ext }`
    : `${ songName }${ ext }`;

  const bookDir = path.join( config.songsDirectory, bookName );
  const targetPath = path.join( bookDir, filename );

  // Resolve template path
  const templatePath = path.join( config.songsDirectory, config.songTemplate );
  try {
    await fs.promises.access( templatePath );
  } catch {
    res.status( 500 ).json({ error: `Song template not found: ${ config.songTemplate }` });
    return;
  }

  // Check target doesn't already exist
  try {
    await fs.promises.access( targetPath );
    res.status( 409 ).json({ error: `File already exists: ${ filename }` });
    return;
  } catch {
    // Good — file doesn't exist
  }

  try {
    // Ensure book directory exists
    await fs.promises.mkdir( bookDir, { recursive: true } );

    // Copy template to target
    await fs.promises.copyFile( templatePath, targetPath );

    // Register in MySQL songs table
    try {
      const relPath = path.join( bookName, filename );
      await findOrCreateSong( songName, songNumber || undefined, bookName, relPath );
    } catch {
      // DB not available — file was created on disk, which is the primary goal
    }

    res.json({
      success: true,
      song: { name: songName, number: songNumber || undefined, book: bookName, ccli: ccli || undefined },
    });
  } catch ( err: any ) {
    console.error( '[SONGS] Error creating song:', err );
    res.status( 500 ).json({ error: 'Failed to create song file' });
  }
});


/**
 * Get a setting value by key.
 *
 * @param key - setting key (e.g., "claude-prompt")
 */
app.get( '/api/settings/:key', async ( req: Request, res: Response ) => {
  try {
    const pool = getPool();
    const [ rows ] = await pool.query<RowDataPacket[]>(
      'SELECT `value`, updated_at FROM settings WHERE `key` = ?',
      [ req.params.key ]
    );
    if ( rows.length === 0 ) {
      res.status( 404 ).json({ error: 'Setting not found' });
      return;
    }
    res.json({ key: req.params.key, value: rows[ 0 ].value, updatedAt: rows[ 0 ].updated_at });
  } catch {
    res.status( 503 ).json({ error: 'Database not available' });
  }
});


/**
 * Update a setting value by key.
 *
 * @param key - setting key
 * @body value - new value (string)
 */
app.put( '/api/settings/:key', async ( req: Request, res: Response ) => {
  const { value } = req.body;
  if ( typeof value !== 'string' ) {
    res.status( 400 ).json({ error: 'value must be a string' });
    return;
  }
  try {
    const pool = getPool();
    await pool.query(
      'INSERT INTO settings (`key`, `value`) VALUES (?, ?) ON DUPLICATE KEY UPDATE `value` = VALUES(`value`)',
      [ req.params.key, value ]
    );
    res.json({ ok: true });
  } catch {
    res.status( 503 ).json({ error: 'Database not available' });
  }
});


/**
 * Backfill song selections from existing Sunday folder shortcuts into the database.
 * Iterates the folder cache and calls syncSelectionsFromFolder for each folder.
 * Idempotent — skips slots that already have active selections.
 */
app.post( '/api/admin/backfill-songs', async ( req: Request, res: Response ) => {
  try {
    if ( !await healthCheck() ) {
      res.status( 503 ).json({ error: 'Database not available' });
      return;
    }

    let foldersScanned = 0;
    let selectionsRecorded = 0;

    for ( const [ date, folder ] of Object.entries( sundayFolders.cache ) ) {
      const songInfos: FolderSongInfo[] = [];

      // Songs
      for ( const [ slotId, song ] of Object.entries( folder.songs ) ) {
        if ( !song ) continue;
        const slotMatch = slotId.match( /^(\d)([a-z])?$/ );
        if ( !slotMatch ) continue;
        songInfos.push({
          slotType: 'song',
          slotNumber: parseInt( slotMatch[ 1 ] ),
          slotSuffix: slotMatch[ 2 ] || undefined,
          name: song.name.replace( /\.(pptx?|ppt)$/i, '' ).replace( /^\d{3}\s*-?\s*/, '' ),
          number: song.number,
          book: song.book,
          target: song.target,
        });
      }

      // Choruses
      for ( let i = 0; i < folder.choruses.length; i++ ) {
        const chorus = folder.choruses[ i ];
        const chorusName = chorus.name.replace( /\.(pptx?|ppt)$/i, '' ).replace( /^\d{3}\s*-?\s*/, '' );
        songInfos.push({
          slotType: 'chorus',
          slotNumber: i + 1,
          name: chorusName,
          book: chorus.book,
          target: chorus.target,
        });
      }

      if ( songInfos.length > 0 ) {
        const before = selectionsRecorded;
        await syncSelectionsFromFolder( date, songInfos, config.songsDirectory );
        selectionsRecorded += songInfos.length;
        foldersScanned++;
      }
    }

    console.log( `[BACKFILL] Scanned ${ foldersScanned } folders, processed ${ selectionsRecorded } song slots` );
    res.json({ ok: true, foldersScanned, selectionsProcessed: selectionsRecorded });
  } catch ( err ) {
    console.error( '[BACKFILL] Error:', err );
    res.status( 500 ).json({ error: 'Backfill failed' });
  }
});


/**
 * function to get all sub folders in a directory recursively
 *
 * @param dir - directory to search
 * @returns array of absolute paths of all sub folders
 *
 * use async yield
 */
async function* getSubFolders( dir: string ): AsyncGenerator<string> {
  const subFolders: string[] = [];
  let entries: fs.Dirent[] = [];
  try {
    entries = await fs.promises.readdir( dir, { withFileTypes: true } );
  } catch ( err ) {
    console.error( `Error reading directory ${ dir }:`, err );
    return;
  }
  for ( const entry of entries ) {
    if ( entry.isDirectory() ) {
      const fullPath = path.join( dir, entry.name );
      subFolders.push( fullPath );
      yield fullPath; // yield the folder path
      // Recursively get subfolders
      for await ( const sub of getSubFolders( fullPath ) ) {
        yield sub; // yield each subfolder found
      }
    }
  }
  return subFolders; // return all found subfolders
}


/**
 * Get all Sundays in the month
 * 
 * @param year - Year (e.g. 2023)
 * @param month - Month (1-12)
 * @return Array of Date objects representing all Sundays in the month
 */
function getSundaysInMonth( year: number, month: number ) {
  const date = new Date( year, month - 1, 1 );

  // Find the first Sunday of the month
  date.setDate( ( 7 - date.getDay() ) % 7 + 1 ); // move to first Sunday of the month
  // must deal with daylight saving time changes by using getTime() + offset
  const offset = date.getTimezoneOffset() * 60 * 1000;

  const sundays: Date[] = [
    new Date( date.getTime() + offset ), // 1-6th day of month
    new Date( date.getTime() + offset +  7 * 86_400_000 ), // 8-14th day of month
    new Date( date.getTime() + offset + 14 * 86_400_000 ), // 15-21st day of month
    new Date( date.getTime() + offset + 21 * 86_400_000 ), // 22-28th day of month
    new Date( date.getTime() + offset + 28 * 86_400_000 )  // 29-31st day of month
  ];

  // Filter out any dates that are not in the requested month
  return sundays.filter( ( d: Date ) => d.getMonth() === month - 1 );
}


/**
 * scan config.outputDirectory for folders and update the cache
 * send an event when a new folder is found
 * 
 */
async function scanFolders() {
  const outputDir = process.env.outputDirectory || process.env.OUTPUT_DIRECTORY || './output';

  // scan existing folders in the cache for changes
  for ( const key of Object.keys( sundayFolders.cache ) ) {
    const folder = sundayFolders.cache[ key ];
    try {
      const stat = await fs.promises.stat( folder.path );
      if ( !stat.isDirectory() ) {
        // remove from cache if it is not a directory
        sundayFolders.delete( key );
      }
    } catch ( error ) {
      console.error( `Error checking folder ${ folder.path }:`, error );
      // remove from cache if it does not exist anymore
      sundayFolders.delete( key );
    }
  }

  // scan for new folders in the output directory
  for await ( const folder of getSubFolders( outputDir ) ) {
    // check if the folder ends with YYYYMMDD format
    const folderName = path.basename( folder );
    const dateMatch = folderName.match( /^\d{4}-?\d{2}-?\d{2}$/ );
    if ( !dateMatch ) {
      continue;
    }

    // Check if the folder is already in the cache
    if ( sundayFolders.get( path.basename( folder ) ) ) {
      continue;
    }

    // Create a new Folder object and add it to the cache
    const newFolder = new SundayFolder( dateMatch[ 0 ], folder );
    sundayFolders.set( dateMatch[ 0 ], newFolder );
  }

  console.log( `Scanned ${ sundayFolders.size } folders in ${ outputDir }` );

  // Auto-detect videos and create ffprobe jobs (only if DB is available)
  try {
    const { healthCheck } = require( './db' );
    if ( !await healthCheck() ) return;

    for ( const key of Object.keys( sundayFolders.cache ) ) {
      const folder = sundayFolders.cache[ key ];
      const vidsDir = path.join( folder.path, 'Vids' );
      const sundayDate = key.replace( /-/g, '' );

      let videoFiles: string[] = [];
      try {
        videoFiles = await fs.promises.readdir( vidsDir );
      } catch {
        continue; // No Vids/ directory
      }

      // Create ffprobe jobs for raw video files (.mp4, .mts)
      // Skip production files — their ffprobe is chained after transcription completes
      const probeFiles = videoFiles.filter( f =>
        /\.(mp4|mts)$/i.test( f ) && !/^\d{8}-production\.mp4$/i.test( f )
      );
      for ( const file of probeFiles ) {
        const inputPath = toJobPath( path.join( vidsDir, file ) );
        const job = await createProbeJobIfNotExists( sundayDate, inputPath );
        if ( job ) {
          io.emit( 'job:created', job );
          console.log( `[JOBS] Auto-created ffprobe job #${ job.id } for ${ file }` );
        }
      }
    }
  } catch ( err ) {
    console.error( '[JOBS] Error during ffprobe auto-detection:', err );
  }

  // Sync song/chorus/closing shortcuts to song_selections table
  try {
    const { healthCheck } = require( './db' );
    if ( !await healthCheck() ) return;

    const songsDir = config.songsDirectory;

    for ( const key of Object.keys( sundayFolders.cache ) ) {
      const folder = sundayFolders.cache[ key ];
      const dateStr = key.replace( /-/g, '' );
      const songs: FolderSongInfo[] = [];

      // Collect songs from the cached SundayFolder
      for ( const [ slotId, song ] of Object.entries( folder.songs ) ) {
        if ( song ) {
          const baseSlot = parseInt( slotId[ 0 ] );
          const suffix = slotId.length > 1 ? slotId[ 1 ] : undefined;
          songs.push({ slotType: 'song', slotNumber: baseSlot, slotSuffix: suffix, name: song.name, number: song.number, book: song.book, target: song.target });
        }
      }

      // Collect choruses
      folder.choruses.forEach( ( c: any, i: number ) => {
        songs.push({ slotType: 'chorus', slotNumber: i + 1, name: c.name, book: c.book, target: c.target });
      });

      // Check for closing song shortcuts by scanning filenames
      try {
        const files = await fs.promises.readdir( folder.path );
        for ( const file of files ) {
          if ( file.toLowerCase().startsWith( 'closing song ' ) && file.toLowerCase().endsWith( '.lnk' ) ) {
            const base = file.replace( /\s*-\s*Shortcut\.lnk$/i, '' );
            const match = base.match( /^closing song\s+(?:(\d+)\s*-\s*)?(.+)$/i );
            if ( match ) {
              songs.push({ slotType: 'closing', slotNumber: 1, name: match[ 2 ].trim(), number: match[ 1 ] || undefined });
            }
          }
        }
      } catch {
        // Folder not readable for closing song scan — skip
      }

      if ( songs.length > 0 ) {
        await syncSelectionsFromFolder( dateStr, songs, songsDir );
      }
    }
  } catch {
    // DB not available — silently skip sync
  }
}


/**
 * Cleanup the single oldest MTS file that is older than 6 months.
 * Runs once per hour. Only deletes one file per invocation for safety.
 */
async function cleanupOldestMts() {
  const oldest = await getOldestMtsVideo();
  if ( !oldest ) return;

  // Resolve relative path to absolute for deletion
  const rootPath = config.rootPath || '%OneDriveConsumer%';
  const absolutePath = resolveToAbsolutePath( path.join( rootPath, oldest.input_path ) );

  // Verify file still exists before deleting
  if ( !fs.existsSync( absolutePath ) ) {
    await deleteVideo( oldest.input_path );
    return;
  }

  try {
    await fs.promises.unlink( absolutePath );
    console.log( `[CLEANUP] Deleted oldest MTS: ${ oldest.filename } (created: ${ oldest.file_created_at })` );

    // Remove video row and cancel ffprobe job
    await deleteVideo( oldest.input_path );
    try {
      if ( await healthCheck() ) {
        const pool = getPool();
        await pool.query(
          `UPDATE jobs SET status = 'cancelled' WHERE type = 'ffprobe' AND input_path = ? AND status IN ('completed', 'pending', 'queued')`,
          [ oldest.input_path ]
        );
      }
    } catch {
      // DB cleanup is best-effort
    }
  } catch ( err ) {
    console.error( `[CLEANUP] Failed to delete MTS: ${ absolutePath }`, err );
  }
}


/**
 * Purge completed, cancelled, and failed jobs older than 90 days.
 * Job logs are cascade-deleted via FK. Parent references are set to NULL.
 */
async function purgeOldJobs() {
  try {
    if ( !await healthCheck() ) return;
    const pool = getPool();
    const [ result ] = await pool.query<ResultSetHeader>(
      `DELETE FROM jobs WHERE status IN ('completed', 'cancelled', 'failed') AND created_at < DATE_SUB(NOW(), INTERVAL 90 DAY)`
    );
    if ( result.affectedRows > 0 ) {
      console.log( `[PURGE] Deleted ${ result.affectedRows } old jobs (>90 days)` );
    }
  } catch ( err ) {
    console.error( '[PURGE] Failed to purge old jobs:', err );
  }
}


/**
 * Returns song/chorus selections for each week in a given month
 */
app.get( '/api/week-song-selections', async ( req: Request, res: Response ) => {
  try {
    const year = parseInt( ( req.query.year as string ) || '' );
    const month = parseInt( ( req.query.month as string ) || '' );
    if ( !year || !month ) {
      res.status( 400 ).json({ error: 'Missing year or month' });
      return;
    }
    const sundays = getSundaysInMonth( year, month );
    const outputDir = process.env.outputDirectory || process.env.OUTPUT_DIRECTORY || './output';
    // For each week, try to find song shortcuts and chorus info
    const result: Record<string, { songs: Record<string, any>; choruses: any[] }> = {};
    for ( let weekIdx = 0; weekIdx < sundays.length; weekIdx++ ) {
      const sunday = sundays[ weekIdx ];
      let weekFolder: string | null = null;
      const subFolderName = `${ sunday.getFullYear() }${ String( sunday.getMonth() + 1 ).padStart( 2, '0' ) }${ String( sunday.getDate() ).padStart( 2, '0' ) }`;

      const possibleFolders = [
        path.join( outputDir, subFolderName ),
        path.join( outputDir, `${ sunday.getFullYear() }`, subFolderName ),
        path.join( outputDir, `${ sunday.getFullYear() } backlog`, subFolderName )
      ]

      // try outputDir/YYYYMMDD
      for ( const folder of possibleFolders ) {
        if ( fs.existsSync( folder ) && fs.statSync( folder ).isDirectory() ) {
          weekFolder = folder;
          break;
        }
      }

      if ( !weekFolder ) continue;

      const songs: Record<string, { number?: string; name: string; shortcut?: string } | null> = { '1': null, '2': null, '3': null };
      let choruses: any[] = [];

      // Find song shortcuts dynamically (supports split slots like "song 2a")
      const files = await fs.promises.readdir( weekFolder );

      // --- Parse songs and choruses, collecting shortcut resolution promises ---
      type PendingResolve = { promise: Promise<string>; apply: ( target: string ) => void };
      const pending: PendingResolve[] = [];

      const songShortcuts = files.filter( ( f: string ) =>
        f.toLowerCase().startsWith( 'song ' ) && f.toLowerCase().endsWith( '.lnk' )
      );

      for ( const shortcut of songShortcuts ) {
        const re = /^song\s+(?<slot>\d+)(?<suffix>[a-z])?\s+(?:(?<number>\d{3})(?:\s*-\s*|\s+))?(?<name>.+?)(?:\s*-\s*Shortcut)?\.lnk$/i;
        const match = shortcut.match( re );
        if ( match && match.groups ) {
          const slotId = match.groups.suffix
            ? `${ match.groups.slot }${ match.groups.suffix.toLowerCase() }`
            : match.groups.slot;
          let name = match.groups.name.trim();
          name = name.replace( /\.[a-zA-Z0-9]{2,5}$/i, '' ).replace( / - Shortcut$/i, '' ).trim();
          const songObj: { number?: string; name: string; shortcut: string } = match.groups.number
            ? { number: match.groups.number, name, shortcut: '' }
            : { name, shortcut: '' };
          songs[ slotId ] = songObj;
          pending.push({
            promise: getRelativeTarget( path.join( weekFolder!, shortcut ), shortcut ),
            apply: ( target ) => { songObj.shortcut = target; },
          });
        }
      }

      // Infer split state: if we see suffix 'a' for base N, ensure at least 'b' key exists
      for ( const slotId of Object.keys( songs ) ) {
        if ( slotId.length === 2 ) {
          const base = slotId[ 0 ];
          // Remove the unsplit base key if suffixed keys exist
          if ( songs[ base ] === null ) delete songs[ base ];
          // Ensure at least 'a' and 'b' exist
          if ( !songs[ `${ base }a` ] && songs[ `${ base }a` ] !== null ) songs[ `${ base }a` ] = null;
          if ( !( `${ base }b` in songs ) ) songs[ `${ base }b` ] = null;
        }
      }

      // --- CHORUS SHORTCUTS ---
      for ( let c = 1; c <= 10; c++ ) {
        const chorusShortcut = files.find( ( f: string ) => f.toLowerCase().startsWith( `chorus ${ c } ` ) && f.toLowerCase().endsWith( '.lnk' ) );
        if ( chorusShortcut ) {
          const re = /^chorus\s+(?<slot>\d+)\s+(?:(?<number>\d{1,5})(?:\s*-\s*|\s+))?(?<name>.+?)(?:\s*-\s*Shortcut)?\.lnk$/i;
          const match = chorusShortcut.match( re );
          let chorusObj: any;
          if ( match && match.groups ) {
            let name = match.groups.name.trim();
            name = name.replace( /\.[a-zA-Z0-9]{2,5}$/i, '' ).replace( / - Shortcut$/i, '' ).trim();
            chorusObj = { name, shortcut: '' };
            if ( match.groups.number ) chorusObj.number = match.groups.number;
          } else {
            let rest = chorusShortcut.replace( /^chorus \d+ /i, '' ).replace( /\.lnk$/i, '' );
            rest = rest.replace( / - Shortcut$/i, '' ).trim();
            chorusObj = { name: rest, shortcut: '' };
          }
          choruses.push( chorusObj );
          pending.push({
            promise: getRelativeTarget( path.join( weekFolder, chorusShortcut ), chorusShortcut ),
            apply: ( target ) => { chorusObj.shortcut = target; },
          });
        }
      }

      // Resolve all shortcut targets in parallel
      const targets = await Promise.all( pending.map( p => p.promise ) );
      targets.forEach( ( target, idx ) => pending[ idx ].apply( target ) );
      // Try to load choruses.json if present
      const chorusJson = path.join( weekFolder, 'choruses.json' );
      if ( fs.existsSync( chorusJson ) ) {
        try {
          const chorusData = JSON.parse( fs.readFileSync( chorusJson, 'utf8' ) );
          if ( Array.isArray( chorusData ) ) choruses = chorusData;
        } catch {}
      }
      result[ String( weekIdx ) ] = { songs, choruses };
    }

    // Enrich songs and choruses with CCLI from the database
    try {
      const pool = getPool();
      // Collect all unique normalized names across all weeks
      const nameSet = new Set<string>();
      for ( const week of Object.values( result ) ) {
        for ( const song of Object.values( week.songs ) ) {
          if ( song?.name ) nameSet.add( normalizeSongName( song.name ) );
        }
        for ( const chorus of week.choruses ) {
          if ( chorus?.name ) nameSet.add( normalizeSongName( chorus.name ) );
        }
      }

      if ( nameSet.size > 0 ) {
        const names = Array.from( nameSet );
        const placeholders = names.map( () => '?' ).join( ', ' );
        const [ rows ] = await pool.query<RowDataPacket[]>(
          `SELECT id, normalized_name, ccli, book FROM songs WHERE normalized_name IN (${ placeholders })`,
          names
        );
        const songDbMap = new Map<string, { id: number; ccli?: string; book?: string }>();
        for ( const row of rows ) {
          songDbMap.set( row.normalized_name, { id: row.id, ccli: row.ccli || undefined, book: row.book || undefined } );
        }

        // Attach id, ccli, and book to each song/chorus
        for ( const week of Object.values( result ) ) {
          for ( const song of Object.values( week.songs ) ) {
            if ( song?.name ) {
              const dbInfo = songDbMap.get( normalizeSongName( song.name ) );
              if ( dbInfo ) {
                song.id = dbInfo.id;
                if ( dbInfo.ccli ) song.ccli = dbInfo.ccli;
                if ( dbInfo.book ) song.book = dbInfo.book;
              }
            }
          }
          for ( const chorus of week.choruses ) {
            if ( chorus?.name ) {
              const dbInfo = songDbMap.get( normalizeSongName( chorus.name ) );
              if ( dbInfo ) {
                chorus.id = dbInfo.id;
                if ( dbInfo.ccli ) chorus.ccli = dbInfo.ccli;
                if ( dbInfo.book ) chorus.book = dbInfo.book;
              }
            }
          }
        }
      }
    } catch {
      // DB unavailable — songs still returned without CCLI
    }

    res.json( result );
  } catch ( e ) {
    res.status( 500 ).json({ error: 'Failed to load week selections' });
  }
});


// --- SHARED HELPERS FOR SONG/CHORUS SHORTCUTS ---

/**
 * Find a song file in the songs directory by matching name and optional number
 *
 * @param song - song object with name and optional number properties
 * @param songsDirectory - root songs directory to search
 * @returns absolute path to the matching song file, or null if not found
 */
function findSongFile( song: any, songsDirectory: string ): string | null {
  /**
   * Recursively walk a directory to find a matching song file
   *
   * @param dir - directory to search
   * @returns absolute path to matching file, or null
   */
  const walk = ( dir: string ): string | null => {
    const files = fs.readdirSync( dir );
    for ( const f of files ) {
      const full = path.join( dir, f );
      if ( fs.statSync( full ).isDirectory() ) {
        const found = walk( full );
        if ( found ) return found;
      } else {
        let base = f.replace( /\.[a-zA-Z0-9]{2,5}$/i, '' ).replace( / - Shortcut$/i, '' ).trim();
        if ( song.number ) {
          if ( f.startsWith( song.number ) && base.includes( song.name ) ) return full;
        } else {
          if ( base === song.name ) return full;
        }
      }
    }
    return null;
  };
  return walk( songsDirectory );
}


/**
 * Extract the book/folder name from a resolved song file path.
 * The book is the immediate parent directory relative to the songs directory.
 *
 * @param songPath - absolute path to the song file
 * @param songsDirectory - root songs directory
 * @returns book name, or undefined if not resolvable
 */
function extractBookFromPath( songPath: string, songsDirectory: string ): string | undefined {
  const normalSong = songPath.replace( /\\/g, '/' );
  const normalBase = songsDirectory.replace( /\\/g, '/' ).replace( /\/$/, '' );
  if ( !normalSong.toLowerCase().startsWith( normalBase.toLowerCase() ) ) return undefined;
  const relative = normalSong.slice( normalBase.length ).replace( /^\//, '' );
  const parts = relative.split( '/' );
  return parts.length >= 2 ? parts[ 0 ] : undefined;
}


/**
 * Delete all shortcut files in a folder that start with the given prefix
 *
 * @param folder - absolute path to the folder
 * @param prefix - prefix to match (e.g. "song 1 ")
 */
function deleteShortcutsByPrefix( folder: string, prefix: string ) {
  const files = fs.readdirSync( folder );
  const shortcuts = files.filter( f => f.toLowerCase().startsWith( prefix ) && f.toLowerCase().endsWith( '.lnk' ) );
  for ( const shortcut of shortcuts ) {
    fs.unlinkSync( path.join( folder, shortcut ) );
    console.log( `[SHORTCUT] Deleted shortcut '${ shortcut }' in folder ${ folder }` );
  }
}


/**
 * API endpoint to update a song selection for a specific date and slot
 */
app.post( '/api/update-song', ( req: Request, res: Response ) => {
  ( async () => {
    try {
      const { weekIdx, slot, song, date } = req.body || {};
      if ( !date || typeof date !== 'string' || !/^\d{8}$/.test( date ) ) {
        return res.status( 400 ).json({ error: 'Missing or invalid date (YYYYMMDD required)' });
      }
      const weekFolder = await resolveFolderPath( date ) || path.join( config.outputDirectory, date );
      if ( !fs.existsSync( weekFolder ) ) fs.mkdirSync( weekFolder, { recursive: true } );
      // Validate slot: accepts "1", "2", "3", "2a", "2b", etc.
      const slotStr = String( slot );
      const slotMatch = slotStr.match( /^(\d)([a-z])?$/ );
      if ( slotMatch ) {
        const slotId = slotMatch[ 2 ] ? `${ slotMatch[ 1 ] }${ slotMatch[ 2 ] }` : slotMatch[ 1 ];
        deleteShortcutsByPrefix( weekFolder, `song ${ slotId } ` );
        const songFile = findSongFile( song, config.songsDirectory );
        if ( !songFile ) {
          console.log( `[SONG SHORTCUT] Song ${ slotId }: Song file not found in library for ${ song.number || '' } ${ song.name }` );
          return res.status( 404 ).json({ error: 'Song file not found in library', status: 'not-found' });
        }
        const shortcutName = `song ${ slotId } ${ song.number ? song.number + ' - ' : '' }${ song.name }`.replace( /[\\/:*?"<>|]/g, '_' ) + ' - Shortcut.lnk';
        const shortcutPath = path.join( weekFolder, shortcutName );
        const desc = `Sunday Song ${ slotId }: ${ song.number ? song.number + ' - ' : '' }${ song.name }`;
        const rootPath = config.rootPath;
        const created = await createShortcut( shortcutPath, desc, songFile, rootPath );
        let status = 'created';
        if ( !created ) {
          status = 'error';
          console.error( `[SONG SHORTCUT] Song ${ slotId }: Failed to create shortcut '${ shortcutName }' in folder ${ weekFolder }` );
        } else {
          console.log( `[SONG SHORTCUT] Song ${ slotId }: Created shortcut '${ shortcutName }' in folder ${ weekFolder }` );
        }

        // Record selection in DB
        try {
          if ( await healthCheck() ) {
            const slotNum = parseInt( slotMatch[ 1 ] );
            const slotSuffix = slotMatch[ 2 ]?.toLowerCase() || undefined;
            const book = song.book || extractBookFromPath( songFile, config.songsDirectory );
            const songId = song.id || await findOrCreateSong( song.name, song.number, book );
            await recordSongSelection( songId, date, 'song', slotNum, slotSuffix );
            console.log( `[SONG DB] Recorded song selection: ${ song.name } (id=${ songId }) → ${ date } slot ${ slotId }` );
          }
        } catch ( err ) {
          console.error( '[SONG DB] Failed to record song selection:', err );
        }

        return res.json({ success: true, status });
      }
      return res.json({ success: true });
    } catch ( e ) {
      console.error( e );
      res.status( 500 ).json({ error: 'Failed to save song selection', status: 'error' });
    }
  })();
});


/**
 * API endpoint to split or unsplit a song slot.
 * Split renames "song N ..." to "song Na ..." and creates empty "Nb" slot.
 * Unsplit renames "song Na ..." back to "song N ..." and deletes other sub-slot shortcuts.
 */
app.post( '/api/song-slot-split', ( req: Request, res: Response ) => {
  ( async () => {
    try {
      const { date, baseSlot, action } = req.body || {};
      if ( !date || typeof date !== 'string' || !/^\d{8}$/.test( date ) ) {
        return res.status( 400 ).json({ error: 'Missing or invalid date (YYYYMMDD required)' });
      }
      const base = parseInt( baseSlot );
      if ( ![ 1, 2, 3 ].includes( base ) ) {
        return res.status( 400 ).json({ error: 'baseSlot must be 1, 2, or 3' });
      }
      if ( ![ 'split', 'unsplit' ].includes( action ) ) {
        return res.status( 400 ).json({ error: 'action must be "split" or "unsplit"' });
      }

      const weekFolder = await resolveFolderPath( date ) || path.join( config.outputDirectory, date );
      if ( !fs.existsSync( weekFolder ) ) {
        return res.status( 404 ).json({ error: 'Folder not found' });
      }

      const files = fs.readdirSync( weekFolder );

      if ( action === 'split' ) {
        // Find existing "song N ..." shortcut and rename to "song Na ..."
        const existing = files.find( f =>
          f.toLowerCase().startsWith( `song ${ base } ` ) && f.toLowerCase().endsWith( '.lnk' )
        );
        if ( existing ) {
          const newName = existing.replace(
            new RegExp( `^(song\\s+${ base })(\\s)`, 'i' ),
            `$1a$2`
          );
          fs.renameSync(
            path.join( weekFolder, existing ),
            path.join( weekFolder, newName )
          );
          console.log( `[SPLIT] Renamed '${ existing }' to '${ newName }'` );
        }
        io.emit( 'folder-changes' );
        return res.json({ success: true, action: 'split', baseSlot: base });
      }

      if ( action === 'unsplit' ) {
        // Find "song Na ..." shortcut and rename to "song N ..."
        const firstSub = files.find( f =>
          f.toLowerCase().startsWith( `song ${ base }a ` ) && f.toLowerCase().endsWith( '.lnk' )
        );
        if ( firstSub ) {
          const newName = firstSub.replace(
            new RegExp( `^(song\\s+${ base })a(\\s)`, 'i' ),
            `$1$2`
          );
          fs.renameSync(
            path.join( weekFolder, firstSub ),
            path.join( weekFolder, newName )
          );
          console.log( `[UNSPLIT] Renamed '${ firstSub }' to '${ newName }'` );
        }
        // Delete all other sub-slot shortcuts (b, c, d, ...)
        const otherSubs = files.filter( f => {
          const lower = f.toLowerCase();
          return lower.endsWith( '.lnk' ) &&
            /^song\s+\d+[b-z]\s/i.test( f ) &&
            f.match( new RegExp( `^song\\s+${ base }[b-z]\\s`, 'i' ) );
        });
        for ( const sub of otherSubs ) {
          fs.unlinkSync( path.join( weekFolder, sub ) );
          console.log( `[UNSPLIT] Deleted '${ sub }'` );
        }
        io.emit( 'folder-changes' );
        return res.json({ success: true, action: 'unsplit', baseSlot: base });
      }
    } catch ( e ) {
      console.error( e );
      res.status( 500 ).json({ error: 'Failed to split/unsplit song slot' });
    }
  })();
});


/**
 * API endpoint to update chorus selections for a specific date
 */
app.post( '/api/chorus-update', ( req: Request, res: Response ) => {
  ( async () => {
    try {
      const { weekIdx, slot, choruses, date } = req.body || {};
      if ( !date || typeof date !== 'string' || !/^\d{8}$/.test( date ) ) {
        return res.status( 400 ).json({ error: 'Missing or invalid date (YYYYMMDD required)' });
      }
      const weekFolder = await resolveFolderPath( date ) || path.join( config.outputDirectory, date );
      if ( !fs.existsSync( weekFolder ) ) fs.mkdirSync( weekFolder, { recursive: true } );
      if ( slot === 'choruses' && Array.isArray( choruses ) ) {
        // Delete all existing chorus shortcuts and log each deletion
        const deletedChoruses: string[] = [];
        const files = fs.readdirSync( weekFolder );
        const chorusShortcuts = files.filter( f => f.toLowerCase().startsWith( 'chorus ' ) && f.toLowerCase().endsWith( '.lnk' ) );
        for ( const shortcut of chorusShortcuts ) {
          fs.unlinkSync( path.join( weekFolder, shortcut ) );
          console.log( `[SHORTCUT] Deleted shortcut '${ shortcut }' in folder ${ weekFolder }` );
          console.log( `[CHORUS SHORTCUT] Deleted shortcut '${ shortcut }' in folder ${ weekFolder }` );
          deletedChoruses.push( shortcut );
        }
        let allCreated = true;
        for ( let i = 0; i < choruses.length; i++ ) {
          const chorus = choruses[ i ];
          const slotNum = i + 1;
          const chorusFile = findSongFile( chorus, config.songsDirectory );
          if ( !chorusFile ) {
            console.log( `[CHORUS SHORTCUT] Chorus ${ slotNum }: File not found in library for ${ chorus.number || '' } ${ chorus.name }` );
            allCreated = false;
            continue;
          }
          const shortcutName = `chorus ${ slotNum } ${ chorus.number ? chorus.number + ' - ' : '' }${ chorus.name }`.replace( /[\\/:*?"<>|]/g, '_' ) + ' - Shortcut.lnk';
          const shortcutPath = path.join( weekFolder, shortcutName );
          const desc = `Sunday Chorus ${ slotNum }: ${ chorus.number ? chorus.number + ' - ' : '' }${ chorus.name }`;
          const rootPath = config.rootPath;
          const created = await createShortcut( shortcutPath, desc, chorusFile, rootPath );
          if ( created ) {
            console.log( `[CHORUS SHORTCUT] Created shortcut '${ shortcutName }' in folder ${ weekFolder }` );
          } else {
            allCreated = false;
            console.error( `[CHORUS SHORTCUT] Failed to create shortcut '${ shortcutName }' in folder ${ weekFolder }` );
          }
        }

        // Record chorus selections in DB
        try {
          if ( await healthCheck() ) {
            await removeAllChorusSelections( date );
            for ( let i = 0; i < choruses.length; i++ ) {
              const chorus = choruses[ i ];
              if ( !chorus.name ) continue;
              const chorusFile = findSongFile( chorus, config.songsDirectory );
              const book = chorus.book || ( chorusFile ? extractBookFromPath( chorusFile, config.songsDirectory ) : undefined );
              const songId = chorus.id || await findOrCreateSong( chorus.name, chorus.number, book );
              await recordSongSelection( songId, date, 'chorus', i + 1 );
            }
            console.log( `[SONG DB] Recorded ${ choruses.length } chorus selections for ${ date }` );
          }
        } catch ( err ) {
          console.error( '[SONG DB] Failed to record chorus selections:', err );
        }

        return res.json({ success: true, status: allCreated ? 'created' : 'partial-error' });
      }
      return res.json({ success: true });
    } catch ( e ) {
      console.error( e );
      res.status( 500 ).json({ error: 'Failed to save chorus selection', status: 'error' });
    }
  })();
});

// Track server start time
const serverStartTime = Date.now();

/**
 * API endpoint to get the last modified date of the newest file in the webapp folder (for hot reload in dev)
 */
app.get( '/api/webapp-last-modified', async ( req: Request, res: Response ) => {
  const webappDir = path.join( __dirname, 'webapp' );
  let latest = 0;

  /**
   * Recursively walk a directory and track the most recent modification time
   *
   * @param dir - directory to walk
   */
  async function walk( dir: string ) {
    let files: string[] = [];
    try {
      files = await fs.promises.readdir( dir );
    } catch { return; }
    for ( const f of files ) {
      const fullPath = path.join( dir, f );
      let stat;
      try {
        stat = await fs.promises.stat( fullPath );
      } catch { continue; }
      if ( stat.isDirectory() ) {
        await walk( fullPath );
      } else {
        if ( stat.mtimeMs > latest ) latest = stat.mtimeMs;
      }
    }
  }

  await walk( webappDir );
  // Use serverStartTime if it is newer than any file
  if ( serverStartTime > latest ) latest = serverStartTime;
  if ( latest > 0 ) {
    res.json({ lastModified: latest, serverStarted: serverStartTime });
  } else {
    res.status( 404 ).json({ error: 'No files found', serverStarted: serverStartTime });
  }
});

// Determine static directory based on environment
const staticDir = process.env.NODE_ENV === 'development'
  ? path.join( __dirname, 'webapp' )
  : path.join( __dirname, '..', 'lib', 'src', 'webapp' );

// Serve book icon images from SONGS_DIRECTORY/icons/
app.use( '/api/books/icons', express.static( path.join( config.songsDirectory, 'icons' ) ) );

// Serve index.html for root
app.get( '/', ( request: Request, response: Response ) => {
  response.sendFile( path.join( staticDir, 'index-old.html' ) );
});

app.use( express.static( staticDir ) );

// Serve Bootstrap CSS and JS from node_modules in development
if ( process.env.NODE_ENV === 'development' ) {
  app.use( '/bootstrap/css', express.static( path.join( __dirname, '..', 'node_modules', 'bootstrap', 'dist', 'css' ) ) );
  app.use( '/bootstrap/js', express.static( path.join( __dirname, '..', 'node_modules', 'bootstrap', 'dist', 'js' ) ) );
}


/**
 * Schedule folder creation on the 2nd Monday of the next month
 * This function checks if today is the 2nd Monday of the next month,
 */
async function scheduleFolderCreation() {
  const now = new Date();
  // Find the 2nd Monday of the next month
  let year = now.getFullYear();
  let month = now.getMonth() + 2; // JS months are 0-based, so +2 for next month
  if ( month > 12 ) {
    month = 1;
    year++;
  }
  // Find the date of the 2nd Monday
  const firstDay = new Date( year, month - 1, 1 );
  let firstMonday = 1 + ( ( 8 - firstDay.getDay() ) % 7 );
  let secondMonday = firstMonday + 7;
  // Only run if today is the 2nd Monday of the month
  const today = new Date();
  if ( today.getFullYear() === year && today.getMonth() + 1 === month && today.getDate() === secondMonday ) {
    const result = await runFolderCreation( { month, year, write: true, foldersOnly: true } );
    if ( !result.success ) {
      console.error( 'Scheduled folder creation failed:', result.error );
    } else {
      console.log( 'Scheduled folder creation succeeded:', result.message );
    }
  }
}


/**
 * Remove date-based shortcuts from the template directory for past Sundays.
 * Deletes any .lnk file matching "YYYY-MM-DD[ TODO].lnk" where the date is today or earlier.
 * Runs every Monday at midnight (after Sunday's service is over).
 */
async function cleanupPastShortcuts() {
  const templateDir = resolveToAbsolutePath( config.templateDirectory );
  const today = new Date();
  today.setHours( 0, 0, 0, 0 );

  let files: string[];
  try {
    files = await fs.promises.readdir( templateDir );
  } catch ( err ) {
    console.error( '[CLEANUP] Cannot read template directory:', err );
    return;
  }

  const datePattern = /^(\d{4}-\d{2}-\d{2})(?:\s+TODO)?\.lnk$/i;
  let deleted = 0;

  for ( const file of files ) {
    const match = file.match( datePattern );
    if ( !match ) continue;

    const fileDate = new Date( match[ 1 ] + 'T00:00:00' );
    if ( isNaN( fileDate.getTime() ) ) continue;

    if ( fileDate <= today ) {
      try {
        await fs.promises.unlink( path.join( templateDir, file ) );
        console.log( `[CLEANUP] Removed past shortcut: ${ file }` );
        deleted++;
      } catch ( err ) {
        console.error( `[CLEANUP] Failed to remove shortcut: ${ file }`, err );
      }
    }
  }

  if ( deleted > 0 ) {
    console.log( `[CLEANUP] Removed ${ deleted } past shortcut(s)` );
  }
}


/**
 * Set up Express endpoints (placeholder for future migration of route handlers)
 */
function setupExpressEndpoints() {

}

// move all global logic and variables down here

setupExpressEndpoints();

// Initialize database and start server
( async () => {
  try {
    await initDb();
    console.log( '[DB] Database initialized successfully' );

    // Seed books table from filesystem directories
    try {
      const pool = getPool();
      const entries = await fs.promises.readdir( config.songsDirectory, { withFileTypes: true } );
      const dirs = entries.filter( e => e.isDirectory() ).map( e => e.name );
      for ( const dir of dirs ) {
        await pool.query(
          'INSERT IGNORE INTO books (name) VALUES (?)',
          [ dir ]
        );
      }
      if ( dirs.length > 0 ) {
        console.log( `[DB] Seeded ${ dirs.length } books from songs directory` );
      }
    } catch ( err ) {
      console.warn( '[DB] Could not seed books table:', err );
    }
  } catch ( err ) {
    console.error( '[DB] Failed to initialize database:', err );
    console.warn( '[DB] Server will continue without job system. Set MYSQL_* env vars to enable.' );
  }

  // Schedule: 2nd Monday of every month at midnight
  schedule.scheduleJob( '0 0 * * 1', scheduleFolderCreation );
  // Cleanup past Sunday shortcuts every Monday at midnight
  schedule.scheduleJob( '0 0 * * 1', cleanupPastShortcuts );
  schedule.scheduleJob( '* * * * *', scanFolders );
  // Recover stale jobs every 2 minutes
  schedule.scheduleJob( '*/2 * * * *', () => recoverStaleJobs( io ) );
  // Cleanup oldest MTS file (>6 months) once per hour
  schedule.scheduleJob( '0 * * * *', cleanupOldestMts );
  // Purge completed/cancelled/failed jobs older than 90 days, daily at 3am
  schedule.scheduleJob( '0 3 * * *', purgeOldJobs );

  const PORT = process?.env?.PORT || 8080;
  server.listen( PORT, () => {
    console.log( `Web server (with socket.io) listening on port ${ PORT }` );
  });
})();

// Graceful shutdown
process.on( 'SIGTERM', async () => {
  console.log( 'SIGTERM received, shutting down...' );
  await closeDb();
  process.exit( 0 );
});
process.on( 'SIGINT', async () => {
  console.log( 'SIGINT received, shutting down...' );
  await closeDb();
  process.exit( 0 );
});
