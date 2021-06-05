#!/usr/bin/env node
const yargs       = require('yargs/yargs')
const { hideBin } = require('yargs/helpers')
const argv        = yargs( hideBin( process.argv ) ).argv
const fs          = require('fs-extra')
const os          = require('os')
const path        = require('path')
const moment      = require('moment')
const to          = require('await-to-js')

console.log({ argv })

const SUNDAY_TEMPLATE = 'Sunday Template.pptx';

const currentWorkingDirectory = typeof argv.dir !== 'undefined' ? argv.dir : __dirname

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
  var sundays = [ 8 - ( new Date( m + '/01/' + y ).getDay() ) ];
  for ( var i = sundays[ 0 ] + 7; i < days; i += 7 ) {
    sundays.push( i );
  }
  return sundays;
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
let futureDate = moment( new Date() ).add( 7, 'day' )

let year = moment( futureDate ).year()
let month = moment( futureDate ).month()+1
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
    console.log(`Coping '${ SUNDAY_TEMPLATE }' to '${ file }'`)
    await fs.copyFile( path.join( currentWorkingDirectory, SUNDAY_TEMPLATE ), filePathTodo )
  } else {
    console.log(`'${ file }${ existsTodo ? ' TODO' : '' }.pptx' already exists, not overwriting`)
  }
});
