#!/usr/bin/env node
const yargs       = require('yargs/yargs')
const { hideBin } = require('yargs/helpers')
const argv        = yargs( hideBin( process.argv ) ).argv
const fs          = require('fs-extra')
const os          = require('os')
const path        = require('path')
const moment      = require('moment')
const to          = require('await-to-js')
const shortcut    = require('create-desktop-shortcuts');

const SUNDAY_TEMPLATE = argv?.templateFile || 'Sunday Template.pptx';
const ext = argv?.ext || 'pptx';
const futureDate = moment( new Date() ).add( 7, 'day' )

const templateDirectory = argv?.templateDirectory || __dirname;
const outputDirectory   = argv?.outputDirectory   || __dirname;
const month     = Number( argv?.month ) || moment( futureDate ).month() + 1;
const year      = Number( argv?.year  ) || moment( futureDate ).year();
const writeMode = !!argv?.write;
const helpMode  = !!argv?.help;

console.log( `${ getMonthName( month ) } ${ year }` )
console.log( `Using Directory: ${ templateDirectory }` )
console.log( `Write mode is ${ writeMode ? 'enabled' : 'disabled, use --write to write files' }\r\n` )

const templateFile = path.join( templateDirectory, SUNDAY_TEMPLATE )


/**
 * get a list of sundays in the month provided
 *
 * @param {number} m human readable month number
 * @param {number} y year
 * @returns an array of days
 */
function sundaysInMonth( m, y ) {
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
function getMonthName( month ) {
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
function checkIfFileExists( path ) {
  return new Promise( resolve => {
    fs.stat( path )
       .then( () => { resolve( true  ) })
      .catch( () => { resolve( false ) })
  })
}


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
 * replaces windows environment variables with their values in the path
 * 
 * @param {string} path windows path strings
 * @returns string
 */
function resolveToAbsolutePath( path ) {
  return path.replace(/%([^%]+)%/g, function( _, key ) {
    return process.env[key];
  });
}


/**
 * go through each file and check if it should be created
 */
sundayFiles.map( async file => {
  let _outputDirectory = path.join( outputDirectory, file.replace( /-/gmi, '' ) );
  let templateFilePath = path.join( templateDirectory, SUNDAY_TEMPLATE );
  let shortcutPathTodo = `${ path.join( templateDirectory, file ) } TODO.lnk`;
  let shortcutPathDone = `${ path.join( templateDirectory, file ) }.lnk`;
  let filePath = `${ path.join( _outputDirectory, file ) }.${ ext }`;
  let existsTodo = await checkIfFileExists( resolveToAbsolutePath( shortcutPathTodo ) );
  let existsDone = await checkIfFileExists( resolveToAbsolutePath( shortcutPathDone ) );

  // create the shortcut
  if ( !existsTodo && !existsDone ) {
    if ( writeMode ) {
      console.log(`Coping '${ SUNDAY_TEMPLATE }' to '${ file }'`)

      await fs.ensureDir( resolveToAbsolutePath( _outputDirectory ) );
      await fs.copyFile( resolveToAbsolutePath( templateFilePath ), resolveToAbsolutePath( filePath ) );
      shortcut({ windows: {
        name: `${ file } TODO`,
        filePath: filePath,
        outputPath: templateDirectory,
      } });
    } else {
      console.log(`Will copy '${ SUNDAY_TEMPLATE }' to '${ filePath }'`)
    }
  } else {
    console.log(`'${ file }${ existsTodo ? ' TODO' : '' }' already exists${ writeMode ? ', not overwriting' : '' }`)
  }
});
