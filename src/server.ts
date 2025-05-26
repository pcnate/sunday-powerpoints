import * as path from 'node:path';
import express, { Request, Response } from 'express';
import * as dotenv from 'dotenv';
import fs from 'fs';
import { spawn, ChildProcessWithoutNullStreams } from 'child_process';
const schedule = require('node-schedule');

dotenv.config();

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

// Start Express web server
const app = express();
app.use(express.json());

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
    let hasFile = false;
    let files: string[] = [];
    try {
      files = await fs.promises.readdir(folderPath);
      hasFile = files.some(f => f.toLowerCase().endsWith('.' + config.ext.toLowerCase()));
    } catch {}
    // Check for mp4 and thumbnail in Vids subdir
    const vidsPath = path.join(folderPath, 'Vids');
    let hasMp4 = false;
    let thumbnail = null;
    try {
      const vidsFiles = await fs.promises.readdir(vidsPath);
      hasMp4 = vidsFiles.some(f => f.toLowerCase().endsWith('.mp4'));
      const thumbFile = vidsFiles.find(f => f.toLowerCase().endsWith('.jpeg') || f.toLowerCase().endsWith('.thm'));
      if (thumbFile) {
        thumbnail = `/api/thumbnail/${encodeURIComponent(folder)}`;
      }
    } catch {}
    // Check for notes file
    const hasNotes = hasNotesFile(folder, files);
    return {
      name: folder,
      hasFile,
      hasMp4,
      thumbnail,
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
    const child: ChildProcessWithoutNullStreams = spawn(
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

// Schedule: 2nd Monday of every month at midnight
schedule.scheduleJob('0 0 * * 1', async function() {
  const now = new Date();
  // Find the 2nd Monday of the next month
  let year = now.getFullYear();
  let month = now.getMonth() + 2; // JS months are 0-based, so +2 for next month
  if (month > 12) {
    month = 1;
    year++;
  }
  // Find the date of the 2nd Monday
  const firstDay = new Date(year, month - 1, 1);
  let firstMonday = 1 + ((8 - firstDay.getDay()) % 7);
  let secondMonday = firstMonday + 7;
  // Only run if today is the 2nd Monday of the month
  const today = new Date();
  if (today.getFullYear() === year && today.getMonth() + 1 === month && today.getDate() === secondMonday) {
    const result = await runFolderCreation({ month, year, write: true, foldersOnly: true });
    if (!result.success) {
      console.error('Scheduled folder creation failed:', result.error);
    } else {
      console.log('Scheduled folder creation succeeded:', result.message);
    }
  }
});

const PORT = 80;
app.listen(PORT, () => {
  console.log(`Web server listening on port ${PORT}`);
});

