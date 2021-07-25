#!/usr/bin/env node
const yargs       = require('yargs/yargs')
const { hideBin } = require('yargs/helpers')
const argv        = yargs( hideBin( process.argv ) ).argv
const fs          = require('fs-extra')
const os          = require('os')
const path        = require('path')
const moment      = require('moment')
const to          = require('await-to-js')

const SUNDAY_TEMPLATE = 'Sunday Template.pptx';
const futureDate = moment( new Date() ).add( 7, 'day' )

const currentWorkingDirectory = argv.dir !== null && argv.dir   !== void 0 ? argv.dir : __dirname;
const month     = argv.month !== null && argv.month !== void 0 ? Number( argv.month ) : moment( futureDate ).month() + 1;
const year      = argv.year  !== null && argv.year  !== void 0 ? Number( argv.year  ) : moment( futureDate ).year();
const writeMode = argv.write !== null && argv.write !== void 0;

console.log( `${ getMonthName( month ) } ${ year }` )
console.log( `Using Directory: ${ currentWorkingDirectory }` )
console.log( `Write mode is ${ writeMode ? 'enabled' : 'disabled, using --write to write files' }\r\n` )

const templateFile = path.join( currentWorkingDirectory, SUNDAY_TEMPLATE )

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

// lets get the date a week out so we can work on next weeks powerpoints

let sundayFiles = sundaysInMonth( month, year ).map( sunday => {
  return [
    year,
    month.toString().padStart( 2, '0' ),
    sunday.toString().padStart( 2, '0' ),
  ].join('-')
});

[''].map( async () => {
  let exists = await checkIfFileExists( templateFile )
  if ( !exists ) {
    console.error(`'${ SUNDAY_TEMPLATE }' not found in '${ currentWorkingDirectory }'` )
    process.exit(1)
  }
  console.log(`Found template: '${ currentWorkingDirectory }\\${ SUNDAY_TEMPLATE }'`)
});

sundayFiles.map( async file => {
  let filePathTodo = path.join( currentWorkingDirectory, file ) + ' TODO.pptx'
  let filePathDone = path.join( currentWorkingDirectory, file ) + '.pptx'
  let existsTodo = await checkIfFileExists( filePathTodo )
  let existsDone = await checkIfFileExists( filePathDone )

  if ( !existsTodo && !existsDone ) {
    if ( writeMode ) {
      console.log(`Coping '${ SUNDAY_TEMPLATE }' to '${ file }'`)
      await fs.copyFile( path.join( currentWorkingDirectory, SUNDAY_TEMPLATE ), filePathTodo )
    } else {
      console.log(`Will copy '${ SUNDAY_TEMPLATE }' to '${ file }'`)
    }
  } else {
    console.log(`'${ file }${ existsTodo ? ' TODO' : '' }.pptx' already exists ${ writeMode ? ', not overwriting' : '' }`)
  }
});
