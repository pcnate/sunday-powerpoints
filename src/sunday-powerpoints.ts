import * as fs from 'node:fs';
import * as path from 'node:path';
import * as ws from 'windows-shortcuts';

/**
 * get a list of sundays in the month provided
 *
 * @param {number} m human readable month number
 * @param {number} y year
 * @returns an array of days
 */
export function sundaysInMonth( m: number, y: number ): number[] {
  var days = new Date( y, m, 0 ).getDate();
  let firstSunday = null;
  var sundays = [];

  let dayOfMonth = 1
  while( dayOfMonth <= days ) {
    
    firstSunday = new Date( m + '/' + dayOfMonth + '/' + y )

    if ( firstSunday.getDay() === 0 ) {
      sundays.push( dayOfMonth )
    }
    
    dayOfMonth++
  }
  
  return sundays;
}


/**
 * get the month name using the month number
 * 
 * @param {number} month human readable month number
 * @returns {string} month name
 */
export function getMonthName( month: number ): string {
  const d = new Date();
  d.setMonth( month - 1 );
  const monthName = d.toLocaleString( "default", { month: "long" } );
  return monthName;
}


/**
 * check if the file exists
 *
 * @param {string} path
 * @returns {boolean} whether or not the file was found
 */
export async function checkIfFileExists( path: string ): Promise<boolean> {
  return await fs.promises.stat( path )
      .then( () => true  )
    .catch( () => false )
}


/**
 * replaces windows environment variables with their values in the path
 * 
 * @param {string} path windows path strings
 * @returns string
 */
export function resolveToAbsolutePath( path: string ): string {
  return path.replace(/%([^%]+)%/g, function (_, key) {
    return process.env[key] ?? '';
  });
}


/**
 * fix existing shortcuts in a directory
 * This function will search for all shortcuts in the specified directory
 * and replace the OneDriveConsumer path with the provided rootPath.
 * If rootPath is not provided, it will replace the path with the variable '^%OneDriveConsumer^%'.
 * This is useful for fixing shortcuts that were created with an old path
 * and need to be updated to use the new OneDriveConsumer variable.
 * 
 * @param directory directory to search for shortcuts
 * @param rootPath optional root path to replace OneDriveConsumer with
 * @returns resolves when all shortcuts have been fixed
 */
export async function fixExistingShortcuts( directory: string, rootPath?: string ): Promise<void> {
  let options = {};
  let files: string[] = await fs.promises.readdir( directory, options );
  files = files.filter( file => file.endsWith('.lnk' ) );
  for ( let file of files ) {
    await fixShortCutPath( path.join( directory, file ), rootPath );
  }
}


/**
 * create a shortcut
 * This function creates a Windows shortcut (.lnk file) at the specified filename.
 * It takes a description and a target file path as parameters.
 * The target path is processed to replace the OneDriveConsumer path with a variable or a provided root path.
 * 
 * @param {path} filename path and filename
 * @param {string} description description of the short cut
 * @param {path} target path and target file
 * @returns 
 */
export async function createShortcut( filename: string, description: string, target: string, rootPath?: string ): Promise<boolean> {
  return new Promise( async resolve => {

    target = await replaceOneDriveConsumerPath( target, rootPath );
    target = target.replace( /\\/gmi, '/' );
    target = target.replace( /\%OneDriveConsumer\%/gmi, '^%OneDriveConsumer^%' )

    interface ShortcutOptions {
      target: string;
      runStyle: number;
      desc: string;
    }

    ws.create(
      filename,
      {
        target: target,
        runStyle: ws.NORMAL,
        desc: description,
      } as ShortcutOptions,
      ( err: Error | null ) => {

        if (err) {
          console.error('Error creating shortcut', filename, err);
          resolve(false);
          return false;
        }

        resolve(true);
        return true;
      }
    );
  })
}


/**
 * get the options for a shortcut
 * 
 * @param {path} filePath path of the shortcut
 * @returns {object}
 */
export async function queryOptions( filePath: string ): Promise<object | false> {
  return new Promise( async resolve => {
    ws.query( filePath, ( error: Error | null, _options?: any ) => {
        if ( !!error ) {
          console.error( 'Error querying shortcut options', error );
          resolve( false );
          return;
        }
        resolve( Object.assign( {}, _options ) );
      }
    );
  })
}


/**
 * update the options for a shortcut
 * 
 * @param filePath path of the shartcut to modify
 * @param options windows-shortcut options
 * @returns true if the options were updated, false if there was an error
 */
export async function updateOptions( filePath: string, options: ws.ShortcutOptions ): Promise<boolean> {
  return new Promise( async resolve => {
    ws.edit( filePath, options, ( error: Error | null ) => {
      if ( !!error ) {
        console.error( 'Error editing shortcut', error );
        resolve( false );
        return;
      }
      resolve( true );
      return;
    })
  });
}


/**
 * replace the absolute path of a shortcut with windows variables
 * 
 * @param {string} filePath path to the shortcut
 * @returns {void}
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
 * replaces the OneDriveConsumer folder with the variable in a path
 * 
 * @param {path} path path to replace
 * @returns {path}
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
 * ensure a directory exists and create it if it does not
 * do not throw an error if the directory already exists
 * 
 * @param dir directory to ensure exists
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
 * write a file to the filesystem with the optional content
 * 
 * @param filePath path to the file
 * @param content content to write to the file
 * @returns resolves when the file has been written
 */
export async function writeFile( filePath: string, content?: string ): Promise<boolean> {
  try {
    await fs.promises.writeFile( filePath, content || '', 'utf8' );
    return true;
  } catch (error) {
    console.error(`Error writing file ${filePath}:`, error);
    return false;
  }
}
  

/**
 * copy a file from one location to another
 * 
 * @param source source file path
 * @param destination destination file path
 * @returns resolves when the file has been copied
 */
export async function copyFile( source: string, destination: string ): Promise<boolean> {
  try {
    await fs.promises.copyFile( source, destination );
    return true;
  } catch (error) {
    console.error(`Error copying file from ${ source } to ${ destination }:`, error);
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
}

export async function runSundayPowerpoints(options: SundayPowerpointsOptions) {
  const { month, year, templateDirectory, outputDirectory, SUNDAY_TEMPLATE, ext, writeMode, rootPath } = options;
  console.log(`${getMonthName(month)} ${year}`);
  console.log(`Using Directory: ${templateDirectory}`);
  console.log(`Write mode is ${writeMode ? 'enabled' : 'disabled, use --write to write files'}\r\n`);

  const templateFile = path.join(templateDirectory, SUNDAY_TEMPLATE);

  // checks for the existence of the template file
  const exists = await checkIfFileExists(resolveToAbsolutePath(templateFile));
  if (!exists) {
    console.error(`'${SUNDAY_TEMPLATE}' not found in '${templateDirectory}'`);
    process.exit(1);
  }
  console.log(`Found template: '${templateDirectory}\\${SUNDAY_TEMPLATE}'`);

  // lets get the date a week out so we can work on next weeks powerpoints
  let sundayFiles: string[] = sundaysInMonth(month, year).map(sunday => {
    return [
      year,
      month.toString().padStart(2, '0'),
      sunday.toString().padStart(2, '0'),
    ].join('-');
  });

  // go through each file and check if it should be created
  for (const file of sundayFiles) {
    let _outputDirectory = path.join(outputDirectory, file.replace(/-/gmi, ''));
    let vidsDirectory = path.join(_outputDirectory, 'Vids');
    let archiveDirectory = path.join(_outputDirectory, 'archive');
    let templateFilePath = path.join(templateDirectory, SUNDAY_TEMPLATE);
    let shortcutPathTodo = `${path.join(templateDirectory, file)} TODO.lnk`;
    let shortcutPathDone = `${path.join(templateDirectory, file)}.lnk`;
    let filePath = `${path.join(_outputDirectory, file)}.${ext}`;
    let notesPath = path.join(_outputDirectory, `${file} Notes.txt`);
    let existsTodo = await checkIfFileExists(resolveToAbsolutePath(shortcutPathTodo));
    let existsDone = await checkIfFileExists(resolveToAbsolutePath(shortcutPathDone));

    // create the shortcut
    if (!existsTodo && !existsDone) {
      if (writeMode) {
        console.log(`Coping '${SUNDAY_TEMPLATE}' to '${file}'`);
        await ensureDir(resolveToAbsolutePath(_outputDirectory));
        await ensureDir(resolveToAbsolutePath(vidsDirectory));
        await ensureDir(resolveToAbsolutePath(archiveDirectory));
        await writeFile(resolveToAbsolutePath(notesPath), '\r\nTitle: \r\n\r\n');
        await copyFile(resolveToAbsolutePath(templateFilePath), resolveToAbsolutePath(filePath));
        await fixExistingShortcuts(resolveToAbsolutePath(_outputDirectory), rootPath);
        await createShortcut(path.join(templateDirectory, `${file} TODO.lnk`), `Sunday ${file}`, filePath, rootPath);
      } else {
        console.log(`Will copy '${SUNDAY_TEMPLATE}' to '${filePath}'`);
      }
    } else {
      console.log(`'${file}${existsTodo ? ' TODO' : ''}' already exists${writeMode ? ', not overwriting' : ''}`);
    }
  }
}
