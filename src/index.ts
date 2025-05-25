#!/usr/bin/env node

import yargs from 'yargs';
import { hideBin } from 'yargs/helpers';
import * as path from 'node:path';
import { runSundayPowerpoints, SundayPowerpointsOptions } from './sunday-powerpoints';
import moment from 'moment';

interface CliArgs {
  templateFile?: string;
  ext?: string;
  templateDirectory?: string;
  outputDirectory?: string;
  month?: string | number;
  year?: string | number;
  write?: boolean;
  help?: boolean;
  rootPath?: string;
  foldersOnly?: boolean;
  [key: string]: unknown;
}

// shortcircuit the file if it is being imported
if ( require.main !== module ) {
  console.log( 'This file is not meant to be imported, please run it directly.' );
  process.exit( 1 );
}

// otherwise run the code
( async () => {

  const argv = yargs(hideBin(process.argv))
    .usage('Usage: $0 [options]')
    .options({
      templateFile: { type: 'string', describe: 'Template file name' },
      ext: { type: 'string', describe: 'File extension', default: 'pptx' },
      templateDirectory: { type: 'string', describe: 'Directory containing the template' },
      outputDirectory: { type: 'string', describe: 'Directory for output files' },
      month: { type: 'number', describe: 'Month (1-12)' },
      year: { type: 'number', describe: 'Year' },
      write: { type: 'boolean', describe: 'Actually write files' },
      rootPath: { type: 'string', describe: 'Root path for OneDriveConsumer variable' },
      foldersOnly: { type: 'boolean', describe: 'Only create folders, do not create files or shortcuts' },
    })
    .help('help')
    .alias('help', 'h')
    .parseSync() as CliArgs;

  const SUNDAY_TEMPLATE = argv.templateFile || 'Sunday Template.pptx';
  const ext = argv.ext || 'pptx';
  const futureDate = moment( new Date() ).add( 7, 'day' )

  const templateDirectory: string  = argv.templateDirectory || __dirname;
  const outputDirectory: string    = argv.outputDirectory || __dirname;
  const month: number              = Number( argv.month ) || moment( futureDate ).month() + 1;
  const year: number               = Number( argv.year ) || moment( futureDate ).year();
  const writeMode: boolean         = !!argv.write;
  const helpMode: boolean          = !!argv.help;
  const rootPath: string|undefined = argv.rootPath || undefined;
  const foldersOnly: boolean         = !!argv.foldersOnly;

  if (helpMode) {
    // yargs.showHelp() does not work as expected in this context, so use yargs(hideBin(process.argv)).getHelp() and print it
    yargs(hideBin(process.argv)).showHelp();
    process.exit(0);
  }

  // everything below this point should be in a function that takes an options object
  await runSundayPowerpoints({
    month,
    year,
    templateDirectory,
    outputDirectory,
    SUNDAY_TEMPLATE,
    ext,
    writeMode,
    rootPath,
    foldersOnly
  });
})();
