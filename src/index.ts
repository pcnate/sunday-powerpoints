#!/usr/bin/env node
const yargs       = require('yargs/yargs')
const { hideBin } = require('yargs/helpers')
const argv        = yargs( hideBin( process.argv ) ).argv
const fs          = require('fs-extra')
const os          = require('os')
const path        = require('path')
const moment      = require('moment')
const to          = require('await-to-js')
const ws          = require('windows-shortcuts');


/**
 * get a list of sundays in the month provided
 *
 * @param {number} m human readable month number
 * @param {number} y year
 * @returns an array of days
 */
export function sundaysInMonth ( m: number, y: number ): number[] {
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
export function getMonthName ( month: number ): string {
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
export function checkIfFileExists ( path: string ): Promise<boolean> {
  return new Promise( resolve => {
    fs.stat( path )
       .then( () => { resolve( true  ) })
      .catch( () => { resolve( false ) })
  })
}


/**
 * replaces windows environment variables with their values in the path
 * 
 * @param {string} path windows path strings
 * @returns string
 */
export function resolveToAbsolutePath ( path: string ): string {
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
export async function fixExistingShortcuts ( directory: string, rootPath?: string ): Promise<void> {
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
export async function createShortcut ( filename: string, description: string, target: string, rootPath?: string ): Promise<boolean> {
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
      (err: Error | null) => {

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
export async function queryOptions ( filePath: string ): Promise<object | false> {
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
 * @param {path} filePath path of the shartcut to modify
 * @param {object} options windows-shortcut options
 * @returns {boolean}
 */
export async function updateOptions ( filePath: string, options: object ): Promise<boolean> {
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
export async function fixShortCutPath ( filePath: string, rootPath?: string ): Promise<boolean> {
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
export async function replaceOneDriveConsumerPath ( path: string, rootPath?: string ): Promise<string> {
  if (rootPath) {
    // Replace any existing OneDriveConsumer variable or path with the CLI argument
    return (path || '').replace(/(([A-Z]:\\Users\\\w+|%USERPROFILE%)\\OneDrive|%OneDriveConsumer%)/gmi, rootPath);
  } else {
    // Fallback to old behavior if not provided
    return (path || '').replace(/(([A-Z]:\\Users\\\w+|%USERPROFILE%)\\OneDrive|%OneDriveConsumer%)/gmi, '^%OneDriveConsumer^%');
  }
}


if ( require.main === module ) {
  ( async () => {

    const SUNDAY_TEMPLATE = argv?.templateFile || 'Sunday Template.pptx';
    const ext = argv?.ext || 'pptx';
    const futureDate = moment( new Date() ).add( 7, 'day' )

    const templateDirectory = argv?.templateDirectory || __dirname;
    const outputDirectory = argv?.outputDirectory || __dirname;
    const month = Number( argv?.month ) || moment( futureDate ).month() + 1;
    const year = Number( argv?.year ) || moment( futureDate ).year();
    const writeMode = !!argv?.write;
    const helpMode = !!argv?.help;
    const rootPath = argv?.rootPath || null;

    console.log( `${getMonthName( month )} ${year}` )
    console.log( `Using Directory: ${templateDirectory}` )
    console.log( `Write mode is ${writeMode ? 'enabled' : 'disabled, use --write to write files'}\r\n` )

    const templateFile = path.join( templateDirectory, SUNDAY_TEMPLATE )

    /**
     * lets get the date a week out so we can work on next weeks powerpoints
     */
    let sundayFiles = sundaysInMonth( month, year ).map( sunday => {
      return [
        year,
        month.toString().padStart( 2, '0' ),
        sunday.toString().padStart( 2, '0' ),
      ].join('-')
    });


    /**
     * checks for the existence of the template file
     */
    [''].map( async () => {
      let exists = await checkIfFileExists( resolveToAbsolutePath( templateFile ) )
      if ( !exists ) {
        console.error(`'${ SUNDAY_TEMPLATE }' not found in '${ templateDirectory }'` )
        process.exit(1)
      }
      console.log(`Found template: '${ templateDirectory }\\${ SUNDAY_TEMPLATE }'`)
    });


    /**
     * go through each file and check if it should be created
     */
    sundayFiles.map( async file => {
      let _outputDirectory = path.join( outputDirectory, file.replace( /-/gmi, '' ) );
      let vidsDirectory    = path.join( _outputDirectory, 'Vids' );
      let archiveDirectory = path.join( _outputDirectory, 'archive' );
      let templateFilePath = path.join( templateDirectory, SUNDAY_TEMPLATE );
      let shortcutPathTodo = `${ path.join( templateDirectory, file ) } TODO.lnk`;
      let shortcutPathDone = `${ path.join( templateDirectory, file ) }.lnk`;
      let filePath = `${ path.join( _outputDirectory, file ) }.${ ext }`;
      let notesPath = path.join( _outputDirectory, `${ file } Notes.txt` );
      let existsTodo = await checkIfFileExists( resolveToAbsolutePath( shortcutPathTodo ) );
      let existsDone = await checkIfFileExists( resolveToAbsolutePath( shortcutPathDone ) );

      // create the shortcut
      if ( !existsTodo && !existsDone ) {
        if ( writeMode ) {
          console.log(`Coping '${ SUNDAY_TEMPLATE }' to '${ file }'`)

          await fs.ensureDir( resolveToAbsolutePath( _outputDirectory ) );
          await fs.ensureDir( resolveToAbsolutePath( vidsDirectory ) );
          await fs.ensureDir( resolveToAbsolutePath( archiveDirectory ) );
          await fs.writeFile( resolveToAbsolutePath( notesPath ), '\r\nTitle: \r\n\r\n' );
          await fs.copyFile( resolveToAbsolutePath( templateFilePath ), resolveToAbsolutePath( filePath ) );
          await fixExistingShortcuts( resolveToAbsolutePath( _outputDirectory ), rootPath );
          await createShortcut( path.join( templateDirectory, `${ file } TODO.lnk` ), `Sunday ${ file }`, filePath, rootPath );
        } else {
          console.log(`Will copy '${ SUNDAY_TEMPLATE }' to '${ filePath }'`)
        }
      } else {
        console.log(`'${ file }${ existsTodo ? ' TODO' : '' }' already exists${ writeMode ? ', not overwriting' : '' }`)
      }
    });

  } )()
} else {
  module.exports = {
    sundaysInMonth,
    getMonthName,
    checkIfFileExists,
    resolveToAbsolutePath,
    fixExistingShortcuts,
    fixShortCutPath,
    replaceOneDriveConsumerPath,
  }
}