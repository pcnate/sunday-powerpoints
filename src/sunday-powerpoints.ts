import * as fs from 'node:fs';
import * as path from 'node:path';
import * as ws from 'windows-shortcuts';

/**
 * Get a list of sundays in the month provided.
 *
 * @param m - human readable month number
 * @param y - year
 * @returns an array of days
 */
export function sundaysInMonth( m: number, y: number ): number[] {
  const days = new Date( y, m, 0 ).getDate();
  let firstSunday = null;
  const sundays: number[] = [];

  let dayOfMonth = 1;
  while ( dayOfMonth <= days ) {

    firstSunday = new Date( m + '/' + dayOfMonth + '/' + y );

    if ( firstSunday.getDay() === 0 ) {
      sundays.push( dayOfMonth );
    }

    dayOfMonth++;
  }

  return sundays;
}


/**
 * Get the month name using the month number.
 *
 * @param month - human readable month number
 * @returns month name
 */
export function getMonthName( month: number ): string {
  const d = new Date();
  d.setMonth( month - 1 );
  const monthName = d.toLocaleString( "default", { month: "long" } );
  return monthName;
}


/**
 * Check if the file exists.
 *
 * @param path - path to check
 * @returns whether or not the file was found
 */
export async function checkIfFileExists( path: string ): Promise<boolean> {
  return await fs.promises.stat( path )
    .then( () => true )
    .catch( () => false );
}


/**
 * Replaces windows environment variables with their values in the path.
 *
 * @param path - windows path strings
 * @returns resolved absolute path
 */
export function resolveToAbsolutePath( path: string ): string {
  return path.replace( /%([^%]+)%/g, function ( _, key ) {
    return process.env[ key ] ?? '';
  } );
}


/**
 * Fix existing shortcuts in a directory.
 * This function will search for all shortcuts in the specified directory
 * and replace the OneDriveConsumer path with the provided rootPath.
 * If rootPath is not provided, it will replace the path with the variable '^%OneDriveConsumer^%'.
 * This is useful for fixing shortcuts that were created with an old path
 * and need to be updated to use the new OneDriveConsumer variable.
 *
 * @param directory - directory to search for shortcuts
 * @param rootPath - optional root path to replace OneDriveConsumer with
 * @returns resolves when all shortcuts have been fixed
 */
export async function fixExistingShortcuts( directory: string, rootPath?: string ): Promise<void> {
  const options = {};
  let files: string[] = await fs.promises.readdir( directory, options );
  files = files.filter( file => file.endsWith( '.lnk' ) );
  for ( const file of files ) {
    await fixShortCutPath( path.join( directory, file ), rootPath );
  }
}


/**
 * Create a shortcut.
 * This function creates a Windows shortcut (.lnk file) at the specified filename.
 * It takes a description and a target file path as parameters.
 * The target path is processed to replace the OneDriveConsumer path with a variable or a provided root path.
 *
 * @param filename - path and filename
 * @param description - description of the shortcut
 * @param target - path and target file
 * @param rootPath - optional root path to replace OneDriveConsumer with
 * @returns whether the shortcut was created successfully
 */
export async function createShortcut( filename: string, description: string, target: string, rootPath?: string ): Promise<boolean> {
  return new Promise( async resolve => {
    target = await replaceOneDriveConsumerPath( target, rootPath );
    target = target.replace( /\\/gmi, '/' );
    target = target.replace( /\%OneDriveConsumer\%/gmi, '^%OneDriveConsumer^%' );
    ws.create(
      filename,
      {
        target: target,
        runStyle: 1, // 1 = NORMAL
        desc: description,
      },
      ( err: any ) => {
        if ( err ) {
          console.error( 'Error creating shortcut', filename, err );
          resolve( false );
          return;
        }
        resolve( true );
      }
    );
  } );
}


/**
 * Get the options for a shortcut.
 *
 * @param filePath - path of the shortcut
 * @returns shortcut options or false on error
 */
export async function queryOptions( filePath: string ): Promise<object | false> {
  return new Promise( resolve => {
    ws.query( filePath, ( error: any, _options?: any ) => {
      if ( !!error ) {
        console.error( 'Error querying shortcut options', error );
        resolve( false );
        return;
      }
      resolve( Object.assign( {}, _options ) );
    } );
  } );
}


/**
 * Update the options for a shortcut.
 *
 * @param filePath - path of the shortcut to modify
 * @param options - windows-shortcut options
 * @returns true if the options were updated, false if there was an error
 */
export async function updateOptions( filePath: string, options: any ): Promise<boolean> {
  return new Promise( resolve => {
    ws.edit( filePath, options, ( error: any ) => {
      if ( !!error ) {
        console.error( 'Error editing shortcut', error );
        resolve( false );
        return;
      }
      resolve( true );
    } );
  } );
}


/**
 * Replace the absolute path of a shortcut with windows variables.
 *
 * @param filePath - path to the shortcut
 * @param rootPath - optional root path to replace OneDriveConsumer with
 * @returns whether the shortcut was fixed successfully
 */
export async function fixShortCutPath( filePath: string, rootPath?: string ): Promise<boolean> {
  let options: any = {};
  let success = false;
  let _options = await queryOptions( filePath );
  if ( !!_options )
    options = _options;
  if ( !!options?.target )
    options.target     = await replaceOneDriveConsumerPath( options.target, rootPath );
  if ( !!options?.workingDir )
    options.workingDir = await replaceOneDriveConsumerPath( options.workingDir, rootPath );
  if ( !!options?.expanded )
    delete options.expanded;
  if ( !!options )
    success = await updateOptions( filePath, options ) as boolean;
  return success;
}


/**
 * Replaces the OneDriveConsumer folder with the variable in a path.
 *
 * @param path - path to replace
 * @param rootPath - optional root path to replace OneDriveConsumer with
 * @returns the path with OneDriveConsumer replaced
 */
export async function replaceOneDriveConsumerPath( path: string, rootPath?: string ): Promise<string> {
  if ( rootPath ) {
    // Replace any existing OneDriveConsumer variable or path with the CLI argument
    return ( path || '' ).replace( /(([A-Z]:\\Users\\\w+|%USERPROFILE%)\\OneDrive|%OneDriveConsumer%)/gmi, rootPath );
  } else {
    // Fallback to old behavior if not provided
    return ( path || '' ).replace( /(([A-Z]:\\Users\\\w+|%USERPROFILE%)\\OneDrive|%OneDriveConsumer%)/gmi, '^%OneDriveConsumer^%' );
  }
}


/**
 * Ensure a directory exists and create it if it does not.
 * Do not throw an error if the directory already exists.
 *
 * @param dir - directory to ensure exists
 * @returns true if the directory was created or already exists
 */
export async function ensureDir( dir: string ): Promise<boolean> {
  try {
    await fs.promises.mkdir( dir, { recursive: true } );
    return true;
  } catch {
    return false;
  }
}


/**
 * Write a file to the filesystem with the optional content.
 *
 * @param filePath - path to the file
 * @param content - content to write to the file
 * @returns resolves when the file has been written
 */
export async function writeFile( filePath: string, content?: string ): Promise<boolean> {
  try {
    await fs.promises.writeFile( filePath, content || '', 'utf8' );
    return true;
  } catch ( error ) {
    console.error( `Error writing file ${ filePath }:`, error );
    return false;
  }
}
  

/**
 * Copy a file from one location to another.
 *
 * @param source - source file path
 * @param destination - destination file path
 * @returns resolves when the file has been copied
 */
export async function copyFile( source: string, destination: string ): Promise<boolean> {
  try {
    await fs.promises.copyFile( source, destination );
    return true;
  } catch ( error ) {
    console.error( `Error copying file from ${ source } to ${ destination }:`, error );
    return false;
  }
}


export interface SundayPowerpointsOptions {
  month: number;
  year: number;
  templateDirectory: string;
  outputDirectory: string;
  SUNDAY_TEMPLATE: string;
  ext: string;
  writeMode: boolean;
  rootPath?: string;
  foldersOnly?: boolean;
  overwrite?: boolean;
}

/**
 * Run the Sunday PowerPoints workflow for a given month.
 * Creates folders, copies templates, and creates shortcuts for each Sunday.
 *
 * @param options - configuration options for the workflow
 */
export async function runSundayPowerpoints( options: SundayPowerpointsOptions ) {
  const { month, year, templateDirectory, outputDirectory, SUNDAY_TEMPLATE, ext, writeMode, rootPath, foldersOnly, overwrite } = options;
  console.log( `${ getMonthName( month ) } ${ year }` );
  console.log( `Using Directory: ${ templateDirectory }` );
  console.log( `Write mode is ${ writeMode ? 'enabled' : 'disabled, use --write to write files' }\r\n` );

  const templateFile = path.join( templateDirectory, SUNDAY_TEMPLATE );

  // checks for the existence of the template file
  if ( !foldersOnly ) {
    const exists = await checkIfFileExists( resolveToAbsolutePath( templateFile ) );
    if ( !exists ) {
      const errorMsg = `'${ SUNDAY_TEMPLATE }' not found in '${ templateDirectory }'`;
      console.error( errorMsg );
      process.send?.({ type: 'error', error: errorMsg });
      process.exit( 1 );
    }
    const foundMsg = `Found template: '${ templateDirectory }\\${ SUNDAY_TEMPLATE }'`;
    console.log( foundMsg );
    process.send?.({ type: 'info', message: foundMsg });
  }

  // lets get the date a week out so we can work on next weeks powerpoints
  const sundayFiles: string[] = sundaysInMonth( month, year ).map( sunday => {
    return [
      year,
      month.toString().padStart( 2, '0' ),
      sunday.toString().padStart( 2, '0' ),
    ].join( '-' );
  } );

  // go through each file and check if it should be created
  for ( const file of sundayFiles ) {
    const _outputDirectory = path.join( outputDirectory, file.replace( /-/gmi, '' ) );
    const vidsDirectory = path.join( _outputDirectory, 'Vids' );
    const archiveDirectory = path.join( _outputDirectory, 'archive' );
    const templateFilePath = path.join( templateDirectory, SUNDAY_TEMPLATE );
    const shortcutPathTodo = `${ path.join( templateDirectory, file ) } TODO.lnk`;
    const shortcutPathDone = `${ path.join( templateDirectory, file ) }.lnk`;
    const filePath = `${ path.join( _outputDirectory, file ) }.${ ext }`;
    const notesPath = path.join( _outputDirectory, `${ file } Notes.txt` );
    const existsTodo = await checkIfFileExists( resolveToAbsolutePath( shortcutPathTodo ) );
    const existsDone = await checkIfFileExists( resolveToAbsolutePath( shortcutPathDone ) );

    // create the shortcut
    if ( !existsTodo && !existsDone ) {
      if ( writeMode ) {
        if ( foldersOnly ) {
          // Only create folders, skip notes, copy, shortcut
          await ensureDir( resolveToAbsolutePath( _outputDirectory ) );
          await ensureDir( resolveToAbsolutePath( vidsDirectory ) );
          await ensureDir( resolveToAbsolutePath( archiveDirectory ) );
          const msg = `Created folders for '${ file }'`;
          console.log( msg );
          process.send?.({ type: 'info', message: msg });
        } else {
          const copyingMsg = `Copying '${ SUNDAY_TEMPLATE }' to '${ file }'`;
          console.log( copyingMsg );
          process.send?.({ type: 'info', message: copyingMsg });
          await ensureDir( resolveToAbsolutePath( _outputDirectory ) );
          await ensureDir( resolveToAbsolutePath( vidsDirectory ) );
          await ensureDir( resolveToAbsolutePath( archiveDirectory ) );
          await writeFile( resolveToAbsolutePath( notesPath ), 'Title: \r\n\r\n' );
          await copyFile( resolveToAbsolutePath( templateFilePath ), resolveToAbsolutePath( filePath ) );
          await fixExistingShortcuts( resolveToAbsolutePath( _outputDirectory ), rootPath );
          await createShortcut( path.join( templateDirectory, `${ file } TODO.lnk` ), `Sunday ${ file }`, filePath, rootPath );
        }
      } else {
        if ( !foldersOnly ) {
          const willCopyMsg = `Will copy '${ SUNDAY_TEMPLATE }' to '${ filePath }'`;
          console.log( willCopyMsg );
          process.send?.({ type: 'info', message: willCopyMsg });
        } else {
          const willCreateFoldersMsg = `Will create folders for '${ file }'`;
          console.log( willCreateFoldersMsg );
          process.send?.({ type: 'info', message: willCreateFoldersMsg });
        }
      }
    } else if ( overwrite && !foldersOnly ) {
      // Shortcut exists but overwrite requested — re-copy the template file
      if ( writeMode ) {
        const overwriteMsg = `Overwriting '${ file }.${ ext }' with template`;
        console.log( overwriteMsg );
        process.send?.({ type: 'info', message: overwriteMsg });
        await copyFile( resolveToAbsolutePath( templateFilePath ), resolveToAbsolutePath( filePath ) );
      } else {
        const willOverwriteMsg = `Will overwrite '${ filePath }' with template`;
        console.log( willOverwriteMsg );
        process.send?.({ type: 'info', message: willOverwriteMsg });
      }
    } else {
      const alreadyExistsMsg = `'${ file }${ existsTodo ? ' TODO' : '' }' already exists${ writeMode ? ', not overwriting' : '' }`;
      console.log( alreadyExistsMsg );
      process.send?.({ type: 'info', message: alreadyExistsMsg });
    }
  }
}
