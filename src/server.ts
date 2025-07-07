import path from 'node:path';
import fs from 'node:fs';
import child_process from 'node:child_process';
import express, { Request, Response } from 'express';
import dotenv from 'dotenv';
import ws from 'windows-shortcuts';
const schedule = require('node-schedule');
import { createShortcut } from './sunday-powerpoints';
import { Server as SocketIOServer } from 'socket.io';
import http from 'http';

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
 * @property {Song|null} song1 - Song 1 shortcut, or null if not set
 * @property {Song|null} song2 - Song 2 shortcut, or null if not set
 * @property {Song|null} song3 - Song 3 shortcut, or null if not set
 * @property {Chorus[]} choruses - Array of chorus shortcuts in the folder
 * @property {string|null} hasNotes - path to notes file if it exists, or null if not found
 */
class SundayFolder {
  name: string; // YYYYMMDD
  path: string; // absolute path to the folder which ends in YYYYMMDD
  presentation: string | null; // path to PowerPoint file if it exists, or null if not found
  thumbnail: string | null; // path to thumbnail image, or null if not found
  video: Video[]; // video files
  song1: Song|null;
  song2: Song|null;
  song3: Song|null;
  choruses: Chorus[];
  hasNotes: string|null;

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
    this.song1 = null;
    this.song2 = null;
    this.song3 = null;
    this.choruses = [];
    this.hasNotes = null;

    // Determine if the folder is archived or backlog based on its path
    this.archived = false;
    this.backlog = false;
    if ( this.path.includes( 'Backlog' ) ) {
      this.backlog = true;
    } else
    // regex match path
    if ( this.path.match(/\/\d{4}\/\d{4}-?\d{2}-?\d{2}$/) ) {
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
    if (this.video.some(v => v.path === _path.toString())) return; // already exists
    this.video.push( new Video( _path ) );
  }


  /**
   * Remove a video from the list
   * 
   * @param path - absolute path to the video file
   */
  removeVideo( path: fs.PathLike ): void {
    const index = this.video.findIndex(v => v.path === path.toString());
    if (index !== -1) this.video.splice(index, 1);
  }


  /**
   * Set a song for the given number
   * 
   * @param _num '1 | 2 | 3' - song number to set (1, 2, or 3)
   * @param _target - absolute path to the target song file
   */
  async setSong( _num: 1 | 2 | 3, _target: fs.PathLike ) {
    if ( _num === 1 && !!this.song1 ) await this.song1.delete();
    if ( _num === 2 && !!this.song2 ) await this.song2.delete();
    if ( _num === 3 && !!this.song3 ) await this.song3.delete();

    const song = new Song( `song ${ _num }`, _target );

    if ( _num === 1 ) this.song1 = song;
    else if ( _num === 2 ) this.song2 = song;
    else if ( _num === 3 ) this.song3 = song;
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
    this.choruses = this.choruses.filter(c => c !== chorus);
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

    const choruses: { name: string, path: string, target: string|null }[] = [];

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
        if ( /^\d{4}-?\d{2}-?\d{2}(\s+)?notes\.txt$/i.test( file ) ) {
          if ( this.hasNotes === filePath ) {
            this.hasNotes = filePath;
            changes = true;
          }
          continue; // skip further checks for notes file
        }

        // Check for song shortcuts and get the target file
        if ( file.toLowerCase().startsWith( 'song ' ) ) {
          const songMatch = file.match(/^song (\d+) /i);
          if ( !songMatch ) continue; // not a song shortcut

          const songNumber = parseInt( songMatch[1] );
          const target = await getShortcutTarget( filePath );
          const song = new Song( `song ${ songNumber }`, target || '', false );

          if ( songNumber === 1 ) {
            // check if song1 is different than song
            if ( this.song1 && this.song1.target === song.target ) continue; // already exists
            this.song1 = song;
            changes = true;
            continue;
          } else if ( songNumber === 2 ) {
            if ( this.song2 && this.song2.target === song.target ) continue; // already exists
            this.song2 = song;
            changes = true;
            continue;
          } else if ( songNumber === 3 ) {
            if ( this.song3 && this.song3.target === song.target ) continue; // already exists
            this.song3 = song;
            changes = true;
            continue;
          }
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
          const prefix = chorus.name.match(/^(chorus\s*\d*)/i)?.[0] || 'chorus';
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
        // if the video is a thm or jpeg file then it is the thumbnail
        if ( videoFile.toLocaleLowerCase().endsWith( '.thm' ) ) {
          // rename the file to a jpeg file
          const newPath = videoPath.replace(/\.thm$/i, '.jpeg');
          await fs.promises.rename( videoPath, newPath );
          this.thumbnail = newPath;
          changes = true;
        } else
        
        // add the video file to the list
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
      song1: this.song1 ? this.song1.toJSON() : null,
      song2: this.song2 ? this.song2.toJSON() : null,
      song3: this.song3 ? this.song3.toJSON() : null,
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
        console.error(`Error querying shortcut ${shortcutPath}:`, error);
        resolve( null );
      } else if ( options && options.target ) {
        resolve( options.target );
      } else {
        resolve( null );
      }
    });
  });
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
    this.name = this.path.split('/').pop() || '';
    this.type = this.name.toLowerCase().endsWith('.mp4') ? 'mp4' : 'mts';

    this.camera = 'unknown';

    // TODO add rules for camera name extraction from the path or metadata
    // main camera uses /\d{5}\.(MP4|MTS)/ format
    const cameraMatch = this.name.match(/(\d{5})\.(mp4|mts)$/i);
    if ( cameraMatch ) this.camera = 'main';
    else if (this.name.toLowerCase().includes('camera2'))
      this.camera = 'camera2';
    else if (this.name.toLowerCase().includes('camera3'))
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
    const match = this.name.match(/^(\d{3})\s*-?\s*/);
    this.number = match ? match[1] : undefined;
    // Book/source is the top-level folder (last part of path)
    this.book = this.target.split( path.sep ).slice(-2, -1)[0];
    this.path = path.join(config.outputDirectory, prefix + ' ' + this.name);

    // whether to auto create the shortcut file
    if ( create ) {
      // Create the shortcut file if it does not exist
      const exists = fs.promises.stat(this.path)
      .then(() => true)
      .catch(() => false);
      
      // Create the shortcut path
      if ( !exists ) {
        createShortcut( this.path, prefix + ' ' + this.name, this.target, config.rootPath || '%OneDriveConsumer%' )
          .then(() => console.log(`Created song shortcut: ${this.path}`))
          .catch((error) => console.error(`Failed to create song shortcut ${this.path}:`, error));
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
    } catch (error) {
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
    this.book = this.target.split( path.sep ).slice(-2, -1)[0];

    // Create the path for the chorus file
    this.path = path.join(config.outputDirectory, prefix + ' ' + this.name);

    // If create is false, do not create the shortcut file
    if ( !create ) return;

    // determine if the file already exists
    const exists = fs.promises.stat(this.path)
      .then(() => true)
      .catch(() => false);

    if ( !exists ) {
      // If the file does not exist, create a shortcut to the target file
      createShortcut( this.path, prefix + ' ' + this.name, this.target, config.rootPath || '%OneDriveConsumer%' )
        .then(() => console.log(`Created chorus shortcut: ${this.path}`))
        .catch((error) => console.error(`Failed to create chorus shortcut ${this.path}:`, error));
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
      .then(() => true)
      .catch((error) => {
        console.error(`Failed to rename chorus file ${this.path}:`, error);
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
      await fs.promises.unlink(this.path);
      console.log(`Deleted chorus file: ${this.path}`);
      return true;
    } catch (error) {
      console.error(`Failed to delete chorus file ${this.path}:`, error);
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
  songsDirectory: process.env.SONGS_DIRECTORY || 'P:/songs' // default to P:/songs, fallback to /songs if not set
};

if (!config.templateFile) {
  console.error('TEMPLATE_FILE environment variable is required.');
  process.exit(1);
}

console.log('Resolved server config:', JSON.stringify( config, null, 2 ) );

// Add a script to run server.ts in a development mode
if (process.env.NODE_ENV === 'development') {
  console.log('Server running in development mode.');
  // Add any dev-specific logic here
} else {
  console.log('Server running in production mode.');
}

// Start Express web server with HTTP server for socket.io
const app = express();
app.use(express.json());

const server = http.createServer(app);
const io = new SocketIOServer(server, {
  cors: {
    origin: '*', // Allow all origins (adjust as needed for production)
    methods: ['GET', 'POST']
  }
});

// Socket.io connection handler
io.on('connection', (socket) => {
  console.log('Socket.io client connected:', socket.id);
  socket.on('disconnect', () => {
    console.log('Socket.io client disconnected:', socket.id);
  });
});

// Log all /api requests, except /api/webapp-last-modified
app.use('/api', (req, res, next) => {
  if (req.originalUrl !== '/api/webapp-last-modified') {
    console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl}`);
  }
  next();
});

// Determine static directory based on environment
const staticDir = process.env.NODE_ENV === 'development'
  ? path.join(__dirname, 'webapp')
  : path.join(__dirname, '..', 'lib', 'src', 'webapp');

app.use(express.static(staticDir));

// API endpoint to list folders in outputDirectory
app.get('/api/folders', async (request: Request, response: Response) => {
  const [error, entries]: [unknown, fs.Dirent[] | null] = await fs.promises.readdir(config.outputDirectory, { withFileTypes: true })
    .then((entries: fs.Dirent[]) => [null, entries] as [null, fs.Dirent[]])
    .catch((error: unknown) => [error, null] as [unknown, null]);
  if (error || !entries) {
    response.status(500).json([]);
    return;
  }
  const folders = entries
    .filter((entry: fs.Dirent) => entry.isDirectory())
    .map((entry: fs.Dirent) => entry.name)
    .sort((a: string, b: string) => b.localeCompare(a)); // Newest first

  // For each folder, check for pptx, mp4, thumbnail, and notes file
  const folderData = await Promise.all(folders.map(async (folder) => {
    const folderPath = path.join(config.outputDirectory, folder);
    // Check for file with configured ext in main folder
    let hasPre = false;
    let files: string[] = [];
    try {
      files = await fs.promises.readdir(folderPath);
      hasPre = files.some(f => f.toLowerCase().endsWith('.' + config.ext.toLowerCase()));
    } catch {}
    // Check for mp4 and thumbnail in Vids subdir
    const vidsPath = path.join(folderPath, 'Vids');
    let hasMp4 = false;
    let thumbnail = null;
    let videoFileName = null;
    try {
      const vidsFiles = await fs.promises.readdir(vidsPath);
      hasMp4 = vidsFiles.some(f => f.toLowerCase().endsWith('.mp4'));
      if (hasMp4) {
        videoFileName = vidsFiles.find(f => f.toLowerCase().endsWith('.mp4'));
      }
      const thumbFile = vidsFiles.find(f => f.toLowerCase().endsWith('.jpeg') || f.toLowerCase().endsWith('.thm'));
      if ( thumbFile ) {
        thumbnail = `/api/thumbnail/${encodeURIComponent(folder)}`;
      }
    } catch {}

    // check for song shortcuts in the folder that start with "song 1", "song 2", "song 3"
    let hasSong1 = false, hasSong2 = false, hasSong3 = false;
    files.forEach(f => {
      const lower = f.toLowerCase();
      if (lower.startsWith('song 1 ')) hasSong1 = true;
      if (lower.startsWith('song 2 ')) hasSong2 = true;
      if (lower.startsWith('song 3 ')) hasSong3 = true;
    });

    // Check for notes file
    const hasNotes = hasNotesFile(folder, files);
    return {
      name: folder,
      hasPre,
      hasMp4,
      thumbnail,
      videoFileName,
      hasSong1,
      hasSong2,
      hasSong3,
      hasNotes,
      files, // for debugging or future use, can be removed if not needed
    };
  }));
  response.json(folderData);
});

// Helper to check for notes file in a folder
function hasNotesFile(folderName: string, files: string[]): boolean {
  // Match notes file: YYYYMMDD-notes.txt, YYYY-MM-DD-notes.txt, YYYYMMDD Notes.txt, YYYY-MM-DD Notes.txt (case-insensitive, flexible on dash/space)
  const notesRegex = /^\d{4}-?\d{2}-?\d{2}(\s+)?notes\.txt$/i;
  return files.some(f => notesRegex.test(f));
}

// Endpoint to serve thumbnail images as JPEGs (no query string, auto-detect first .jpeg or .thm)
app.get('/api/thumbnail/:folder', async (req: Request, res: Response) => {
  const folder = req.params.folder;
  if (!folder) {
    res.status(400).send('Missing folder');
    return;
  }
  const vidsPath = path.join(config.outputDirectory, folder, 'Vids');
  let files: string[] = [];
  try {
    files = await fs.promises.readdir(vidsPath);
  } catch {
    res.status(404).send('Vids folder not found');
    return;
  }
  const thumbFile = files.find(f => f.toLowerCase().endsWith('.jpeg') || f.toLowerCase().endsWith('.thm'));
  if (!thumbFile) {
    res.status(404).send('Thumbnail not found');
    return;
  }
  const filePath = path.join(vidsPath, thumbFile);
  fs.promises.readFile(filePath)
    .then(data => {
      res.setHeader('Content-Type', 'image/jpeg');
      res.send(data);
    })
    .catch(() => {
      res.status(404).send('Thumbnail not found');
    });
});

// API endpoint to provide template file details
app.get('/api/template-info', async (req: Request, res: Response) => {
  const templateFile = config.templateFile || '';
  const templatePath = path.join(config.templateDirectory, templateFile);
  function humanFileSize(bytes: number): string {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }
  try {
    const stat = await fs.promises.stat(templatePath);
    res.json({
      name: templateFile,
      lastModified: stat.mtime.toLocaleString(),
      size: humanFileSize(stat.size)
    });
  } catch {
    res.json({ error: 'Template file not found' });
  }
});

// Helper function to run the folder creation script
function runFolderCreation({ month, year, write, foldersOnly }: { month: number, year: number, write: boolean, foldersOnly: boolean }): Promise<{ success: boolean, message?: string, error?: string }> {
  return new Promise((resolve) => {
    let command: string;
    let args: string[];
    let cwd: string;
    let useShell = false;
    function quoteIfNeeded(val: string): string {
      return /\s/.test(val) ? `"${val.replace(/"/g, '\"')}"` : val;
    }
    if (process.env.NODE_ENV === 'development') {
      command = process.platform === 'win32' ? 'npx.cmd' : 'npx';
      args = [
        'ts-node',
        'src/index.ts',
        '--month', String(month),
        '--year', String(year),
        '--templateFile', quoteIfNeeded(String(config.templateFile || '')),
        '--templateDirectory', quoteIfNeeded(String(config.templateDirectory || '')),
        '--outputDirectory', quoteIfNeeded(String(config.outputDirectory || '')),
        '--ext', quoteIfNeeded(String(config.ext || '')),
        '--rootPath', quoteIfNeeded(String(config.rootPath || ''))
      ];
      if (write) args.push('--write');
      if (foldersOnly !== false) args.push('--folders-only');
      cwd = path.resolve(__dirname, '..'); // project root
      useShell = true; // npx on Windows sometimes needs shell
    } else {
      command = process.execPath; // node executable
      args = [
        'lib/index.js',
        '--month', String(month),
        '--year', String(year),
        '--templateFile', quoteIfNeeded(String(config.templateFile || '')),
        '--templateDirectory', quoteIfNeeded(String(config.templateDirectory || '')),
        '--outputDirectory', quoteIfNeeded(String(config.outputDirectory || '')),
        '--ext', quoteIfNeeded(String(config.ext || '')),
        '--rootPath', quoteIfNeeded(String(config.rootPath || ''))
      ];
      if (write) args.push('--write');
      if (foldersOnly !== false) args.push('--folders-only');
      cwd = path.resolve(__dirname, '..'); // project root
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
    child.stdout.on('data', (data: Buffer) => { output += data.toString(); });
    child.stderr.on('data', (data: Buffer) => { error += data.toString(); });

    child.on('close', (code: number) => {
      if (code === 0) {
        resolve({ success: true, message: output.trim() });
      } else {
        resolve({ success: false, error: (error.trim() || output.trim() || 'Script failed.') });
      }
    });
  });
}

// API endpoint to run the folder creation script (folders only)
app.post('/api/run-folders', async (req: Request, res: Response) => {
  const { month, year, write } = req.body || {};
  if (!month || !year) {
    res.status(400).json({ error: 'Month and year are required.' });
    return;
  }
  const result = await runFolderCreation({ month: Number(month), year: Number(year), write: !!write, foldersOnly: true });
  res.json(result);
});

// API endpoint to run the template copy script (full template copy)
app.post('/api/copy-template', async (req: Request, res: Response) => {
  const { month, year, write } = req.body || {};
  if (!month || !year) {
    res.status(400).json({ error: 'Month and year are required.' });
    return;
  }
  const result = await runFolderCreation({ month: Number(month), year: Number(year), write: !!write, foldersOnly: false });
  res.json(result);
});

// API endpoint to get notes file for a folder
app.get('/api/notes', (req, res) => {
  (async () => {
    const folder = req.query.folder as string;
    if (!folder) return res.status(400).send('Missing folder');
    const folderPath = path.join(config.outputDirectory, folder);
    let files: string[] = [];
    try {
      files = await fs.promises.readdir(folderPath);
    } catch {
      return res.status(404).send('Folder not found');
    }
    const notesFile = files.find(f => /^\d{4}-?\d{2}-?\d{2}(\s+)?notes\.txt$/i.test(f));
    if (!notesFile) return res.status(404).send('Notes file not found');
    const notesPath = path.join(folderPath, notesFile);
    try {
      const text = await fs.promises.readFile(notesPath, 'utf8');
      res.type('text/plain').send(text);
    } catch {
      res.status(500).send('Could not read notes file');
    }
  })();
});

// API endpoint to update notes file for a folder
app.post('/api/notes', function(req: any, res: any) {
  const folder = req.query.folder as string;
  if (!folder) return res.status(400).send('Missing folder');
  const folderPath = path.join(config.outputDirectory, folder);
  fs.promises.readdir(folderPath)
    .then(files => {
      const notesFile = files.find(f => /^\d{4}-?\d{2}-?\d{2}(\s+)?notes\.txt$/i.test(f));
      if (!notesFile) { res.status(404).send('Notes file not found'); return; }
      const notesPath = path.join(folderPath, notesFile);
      let body = '';
      req.on('data', (chunk: Buffer) => { body += chunk.toString(); });
      req.on('end', () => {
        fs.promises.writeFile(notesPath, body, 'utf8')
          .then(() => res.send('OK'))
          .catch(() => res.status(500).send('Could not save notes file'));
      });
    })
    .catch(() => res.status(404).send('Folder not found'));
});

// API endpoint to list songs in songsDirectory
app.get('/api/songs', async (req: Request, res: Response) => {
  const songsDir = config.songsDirectory;
  let songFiles: { name: string, number?: string, book?: string, lastModified?: string }[] = [];
  const ext = '.' + (config.ext || 'pptx').toLowerCase();
  async function walk(dir: string, relBase: string = '') {
    let files: string[] = [];
    try {
      files = await fs.promises.readdir(dir);
    } catch (err) {
      return;
    }
    for (const f of files) {
      const fullPath = path.join(dir, f);
      const relPath = relBase ? path.join(relBase, f) : f;
      let stat;
      try {
        stat = await fs.promises.stat(fullPath);
      } catch { continue; }
      if (stat.isDirectory()) {
        await walk(fullPath, relPath);
      } else if (f.toLowerCase().endsWith(ext)) {
        // Extract song number if filename starts with exactly 3 digits, optionally followed by space or dash
        // Example: 123 - Song Name.pptx or 123- Song Name.pptx or 123-Name.pptx
        const match = f.match(/^(\d{3})\s*-?\s*/);
        const number = match ? match[1] : undefined;
        // Book/source is the top-level folder (first part of relPath)
        const book = relPath.split(path.sep)[0];
        // Remove number and extension from name
        let name = f.replace(/^(\d{3})\s*-?\s*/, '').replace(new RegExp(ext + '$', 'i'), '');
        // Format last modified date
        const lastModified = stat.mtime.toLocaleString();
        songFiles.push({ name, number, book, lastModified });
      }
    }
  }
  await walk(songsDir);
  res.json(songFiles);
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
    entries = await fs.promises.readdir(dir, { withFileTypes: true });
  } catch (err) {
    console.error(`Error reading directory ${dir}:`, err);
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
  const sundays: Date[] = [
    new Date( date ), // 1-6th day of month
    new Date( date.getTime() +  7 * 86_400_000 ), // 8-14th day of month
    new Date( date.getTime() + 14 * 86_400_000 ), // 15-21st day of month
    new Date( date.getTime() + 21 * 86_400_000 ), // 22-28th day of month
    new Date( date.getTime() + 28 * 86_400_000 )  // 29-31st day of month
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
    const folder = sundayFolders.cache[key];
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
    const dateMatch = folderName.match(/^\d{4}-?\d{2}-?\d{2}$/);
    if ( !dateMatch ) {
      continue;
    }
    
    // Check if the folder is already in the cache
    if ( sundayFolders.get( path.basename( folder ) ) ) {
      continue;
    }
    
    // Create a new Folder object and add it to the cache
    const newFolder = new SundayFolder( dateMatch[0], folder );
    sundayFolders.set( dateMatch[0], newFolder );
  }

  console.log( `Scanned ${ sundayFolders.size } folders in ${ outputDir }` );
}


// --- SONG SELECTION API ENDPOINTS ---
// Returns song/chorus selections for each week in a given month
app.get('/api/week-song-selections', async (req: Request, res: Response) => {
  try {
    const year = parseInt((req.query.year as string) || '');
    const month = parseInt((req.query.month as string) || '');
    if (!year || !month) {
      res.status(400).json({ error: 'Missing year or month' });
      return;
    }
    const sundays = getSundaysInMonth( year, month );
    const outputDir = process.env.outputDirectory || process.env.OUTPUT_DIRECTORY || './output';
    // For each week, try to find song shortcuts and chorus info
    const result: Record<string, { songs: any[]; choruses: any[] }> = {};
    for ( let weekIdx = 0; weekIdx < sundays.length; weekIdx++ ) {
      const sunday = sundays[weekIdx];
      let weekFolder: string|null = null;
      const subFolderName = `${sunday.getFullYear()}${String( sunday.getMonth() + 1 ).padStart( 2, '0' )}${String( sunday.getDate() ).padStart( 2, '0' )}`;

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

      const songs: (null | { number?: string; name: string })[] = [null, null, null];
      let choruses: any[] = [];

      // Find song shortcuts: song 1, song 2, song 3
      const files = await fs.promises.readdir( weekFolder );
      for (let i = 1; i <= 3; i++) {
        // Regex: song <slot> <number?> <name> [ - Shortcut].lnk (case-insensitive)
        // Examples:
        //   song 1 203 - Tell Me the Story of Jesus - Shortcut.lnk
        //   song 2 203 Tell Me the Story of Jesus.lnk
        //   song 3 Tell Me the Story of Jesus - Shortcut.lnk
        //   song 1 203- Tell Me the Story of Jesus.lnk
        //   song 1 203 -Tell Me the Story of Jesus.lnk
        //   song 1 Tell Me the Story.lnk
        const shortcut = files.find((f: string) => f.toLowerCase().startsWith(`song ${i} `) && f.toLowerCase().endsWith('.lnk'));
        if (shortcut) {
          const re = /^song\s+(?<slot>\d+)\s+(?:(?<number>\d{3})(?:\s*-\s*|\s+))?(?<name>.+?)(?:\s*-\s*Shortcut)?\.lnk$/i;
          const match = shortcut.match(re);
          if (match && match.groups) {
            let name = match.groups.name.trim();
            // Remove any trailing file extension or ' - Shortcut' if present (shouldn't be, but just in case)
            name = name.replace(/\.[a-zA-Z0-9]{2,5}$/i, '').replace(/ - Shortcut$/i, '').trim();
            if (match.groups.number) {
              songs[i-1] = { number: match.groups.number, name };
            } else {
              songs[i-1] = { name };
            }
          } else {
            // fallback: just use the rest of the name
            let rest = shortcut.replace(/^song \d+ /i, '').replace(/\.lnk$/i, '');
            rest = rest.replace(/ - Shortcut$/i, '').trim();
            songs[i-1] = { name: rest };
          }
        }
      }
      // --- CHORUS SHORTCUTS ---
      // Detect chorus shortcuts: chorus 1, chorus 2, etc.
      for (let c = 1; c <= 10; c++) { // support up to 10 choruses per week
        const chorusShortcut = files.find((f: string) => f.toLowerCase().startsWith(`chorus ${c} `) && f.toLowerCase().endsWith('.lnk'));
        if (chorusShortcut) {
          // Regex: chorus <slot> <number?> <name> [ - Shortcut].lnk (case-insensitive)
          // Examples:
          //   chorus 1 10,000 Reasons.pptx - Shortcut.lnk
          //   chorus 2 123 - Name.lnk
          //   chorus 3 Name - Shortcut.lnk
          //   chorus 1 123- Name.lnk
          //   chorus 1 123 -Name.lnk
          //   chorus 1 Name.lnk
          const re = /^chorus\s+(?<slot>\d+)\s+(?:(?<number>\d{1,5})(?:\s*-\s*|\s+))?(?<name>.+?)(?:\s*-\s*Shortcut)?\.lnk$/i;
          const match = chorusShortcut.match(re);
          if (match && match.groups) {
            let name = match.groups.name.trim();
            // Remove any trailing file extension or ' - Shortcut' if present
            name = name.replace(/\.[a-zA-Z0-9]{2,5}$/i, '').replace(/ - Shortcut$/i, '').trim();
            let chorusObj: any = { name };
            if (match.groups.number) chorusObj.number = match.groups.number;
            choruses.push(chorusObj);
          } else {
            // fallback: just use the rest of the name
            let rest = chorusShortcut.replace(/^chorus \d+ /i, '').replace(/\.lnk$/i, '');
            rest = rest.replace(/ - Shortcut$/i, '').trim();
            choruses.push({ name: rest });
          }
        }
      }
      // Try to load choruses.json if present
      const chorusJson = path.join(weekFolder, 'choruses.json');
      if (fs.existsSync(chorusJson)) {
        try {
          const chorusData = JSON.parse(fs.readFileSync(chorusJson, 'utf8'));
          if (Array.isArray(chorusData)) choruses = chorusData;
        } catch {}
      }
      result[String(weekIdx)] = { songs, choruses };
    }
    res.json(result);
  } catch (e) {
    res.status(500).json({ error: 'Failed to load week selections' });
  }
});

// --- SHARED HELPERS FOR SONG/CHORUS SHORTCUTS ---
function findSongFile(song: any, songsDirectory: string): string | null {
  const walk = (dir: string): string | null => {
    const files = fs.readdirSync(dir);
    for (const f of files) {
      const full = path.join(dir, f);
      if (fs.statSync(full).isDirectory()) {
        const found = walk(full);
        if (found) return found;
      } else {
        let base = f.replace(/\.[a-zA-Z0-9]{2,5}$/i, '').replace(/ - Shortcut$/i, '').trim();
        if (song.number) {
          if (f.startsWith(song.number) && base.includes(song.name)) return full;
        } else {
          if (base === song.name) return full;
        }
      }
    }
    return null;
  };
  return walk(songsDirectory);
}

function deleteShortcutsByPrefix(folder: string, prefix: string) {
  const files = fs.readdirSync(folder);
  const shortcuts = files.filter(f => f.toLowerCase().startsWith(prefix) && f.toLowerCase().endsWith('.lnk'));
  for (const shortcut of shortcuts) {
    fs.unlinkSync(path.join(folder, shortcut));
    console.log(`[SHORTCUT] Deleted shortcut '${shortcut}' in folder ${folder}`);
  }
}

// --- SONG SELECTION API ENDPOINT (renamed) ---
app.post('/api/update-song', (req: Request, res: Response) => {
  (async () => {
    try {
      const { weekIdx, slot, song, date } = req.body || {};
      if (!date || typeof date !== 'string' || !/^\d{8}$/.test(date)) {
        return res.status(400).json({ error: 'Missing or invalid date (YYYYMMDD required)' });
      }
      let y = parseInt(date.slice(0, 4));
      let m = parseInt(date.slice(4, 6));
      let d = parseInt(date.slice(6, 8));
      const outputDir = process.env.outputDirectory || process.env.OUTPUT_DIRECTORY || './output';
      const weekFolder = path.join(outputDir, `${y}${String(m).padStart(2,'0')}${String(d).padStart(2,'0')}`);
      if (!fs.existsSync(weekFolder)) fs.mkdirSync(weekFolder, { recursive: true });
      if (slot === 1 || slot === 2 || slot === 3 || slot === '1' || slot === '2' || slot === '3') {
        const slotNum = Number(slot);
        deleteShortcutsByPrefix(weekFolder, `song ${slotNum} `);
        const songFile = findSongFile(song, config.songsDirectory);
        if (!songFile) {
          console.log(`[SONG SHORTCUT] Song ${slotNum}: Song file not found in library for ${song.number || ''} ${song.name}`);
          return res.status(404).json({ error: 'Song file not found in library', status: 'not-found' });
        }
        const shortcutName = `song ${slotNum} ${song.number ? song.number + ' - ' : ''}${song.name}`.replace(/[\\/:*?"<>|]/g, '_') + ' - Shortcut.lnk';
        const shortcutPath = path.join(weekFolder, shortcutName);
        const desc = `Sunday Song ${slotNum}: ${song.number ? song.number + ' - ' : ''}${song.name}`;
        const rootPath = config.rootPath;
        const created = await createShortcut(shortcutPath, desc, songFile, rootPath);
        let status = 'created';
        if (!created) {
          status = 'error';
          console.error(`[SONG SHORTCUT] Song ${slotNum}: Failed to create shortcut '${shortcutName}' in folder ${weekFolder}`);
        } else {
          console.log(`[SONG SHORTCUT] Song ${slotNum}: Created shortcut '${shortcutName}' in folder ${weekFolder}`);
        }
        return res.json({ success: true, status });
      }
      return res.json({ success: true });
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: 'Failed to save song selection', status: 'error' });
    }
  })();
});

// --- CHORUS UPDATE API ENDPOINT ---
app.post('/api/chorus-update', (req: Request, res: Response) => {
  (async () => {
    try {
      const { weekIdx, slot, choruses, date } = req.body || {};
      if (!date || typeof date !== 'string' || !/^\d{8}$/.test(date)) {
        return res.status(400).json({ error: 'Missing or invalid date (YYYYMMDD required)' });
      }
      let y = parseInt(date.slice(0, 4));
      let m = parseInt(date.slice(4, 6));
      let d = parseInt(date.slice(6, 8));
      const outputDir = process.env.outputDirectory || process.env.OUTPUT_DIRECTORY || './output';
      const weekFolder = path.join(outputDir, `${y}${String(m).padStart(2,'0')}${String(d).padStart(2,'0')}`);
      if (!fs.existsSync(weekFolder)) fs.mkdirSync(weekFolder, { recursive: true });
      if (slot === 'choruses' && Array.isArray(choruses)) {
        // Delete all existing chorus shortcuts and log each deletion
        const deletedChoruses: string[] = [];
        const files = fs.readdirSync(weekFolder);
        const chorusShortcuts = files.filter(f => f.toLowerCase().startsWith('chorus ') && f.toLowerCase().endsWith('.lnk'));
        for (const shortcut of chorusShortcuts) {
          fs.unlinkSync(path.join(weekFolder, shortcut));
          console.log(`[SHORTCUT] Deleted shortcut '${shortcut}' in folder ${weekFolder}`);
          console.log(`[CHORUS SHORTCUT] Deleted shortcut '${shortcut}' in folder ${weekFolder}`);
          deletedChoruses.push(shortcut);
        }
        let allCreated = true;
        for (let i = 0; i < choruses.length; i++) {
          const chorus = choruses[i];
          const slotNum = i + 1;
          const chorusFile = findSongFile(chorus, config.songsDirectory);
          if (!chorusFile) {
            console.log(`[CHORUS SHORTCUT] Chorus ${slotNum}: File not found in library for ${chorus.number || ''} ${chorus.name}`);
            allCreated = false;
            continue;
          }
          const shortcutName = `chorus ${slotNum} ${chorus.number ? chorus.number + ' - ' : ''}${chorus.name}`.replace(/[\\/:*?"<>|]/g, '_') + ' - Shortcut.lnk';
          const shortcutPath = path.join(weekFolder, shortcutName);
          const desc = `Sunday Chorus ${slotNum}: ${chorus.number ? chorus.number + ' - ' : ''}${chorus.name}`;
          const rootPath = config.rootPath;
          const created = await createShortcut(shortcutPath, desc, chorusFile, rootPath);
          if (created) {
            console.log(`[CHORUS SHORTCUT] Created shortcut '${shortcutName}' in folder ${weekFolder}`);
          } else {
            allCreated = false;
            console.error(`[CHORUS SHORTCUT] Failed to create shortcut '${shortcutName}' in folder ${weekFolder}`);
          }
        }
        return res.json({ success: true, status: allCreated ? 'created' : 'partial-error' });
      }
      return res.json({ success: true });
    } catch (e) {
      console.error(e);
      res.status(500).json({ error: 'Failed to save chorus selection', status: 'error' });
    }
  })();
});

// Track server start time
const serverStartTime = Date.now();

// API endpoint to get the last modified date of the newest file in the webapp folder (for hot reload in dev)
app.get('/api/webapp-last-modified', async (req: Request, res: Response) => {
  const webappDir = path.join(__dirname, 'webapp');
  let latest = 0;
  async function walk(dir: string) {
    let files: string[] = [];
    try {
      files = await fs.promises.readdir(dir);
    } catch { return; }
    for (const f of files) {
      const fullPath = path.join(dir, f);
      let stat;
      try {
        stat = await fs.promises.stat(fullPath);
      } catch { continue; }
      if (stat.isDirectory()) {
        await walk(fullPath);
      } else {
        if (stat.mtimeMs > latest) latest = stat.mtimeMs;
      }
    }
  }
  await walk(webappDir);
  // Use serverStartTime if it is newer than any file
  if (serverStartTime > latest) latest = serverStartTime;
  if (latest > 0) {
    res.json({ lastModified: latest, serverStarted: serverStartTime });
  } else {
    res.status(404).json({ error: 'No files found', serverStarted: serverStartTime });
  }
});

// Serve index.html for root
app.get('/', (request: Request, response: Response) => {
  response.sendFile(path.join(staticDir, 'index.html'));
});

// Serve Bootstrap CSS and JS from node_modules in development
if (process.env.NODE_ENV === 'development') {
  app.use('/bootstrap/css', express.static(path.join(__dirname, '..', 'node_modules', 'bootstrap', 'dist', 'css')));
  app.use('/bootstrap/js', express.static(path.join(__dirname, '..', 'node_modules', 'bootstrap', 'dist', 'js')));
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

function setupExpressEndpoints() {

}

// move all global logic and variables down here

setupExpressEndpoints();
// Schedule: 2nd Monday of every month at midnight
schedule.scheduleJob('0 0 * * 1', scheduleFolderCreation );
schedule.scheduleJob('* * * * *', scanFolders );

const PORT = process?.env?.PORT || 8080;
server.listen( PORT, () => {
  console.log(`Web server (with socket.io) listening on port ${ PORT }`);
});
