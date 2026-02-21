# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Sunday-powerpoints is a hybrid Node.js/Angular application designed to automate PowerPoint management for weekly church services. The system creates folders for each Sunday of the month, manages song shortcuts, tracks presentation files, and provides a web interface for song selection and file management.

## Development Commands

### Running the Application

- **Development server (Node.js backend with hot-reload)**: `npm run dev`
  - Uses nodemon to watch TypeScript files in `src/` (excluding `src/webapp`)
  - Runs Express server on port specified in `.env` (default: 8080)

- **CLI mode (standalone)**: `npm start -- [options]`
  - Runs the folder/shortcut creation logic without the server
  - See `src/index.ts` for CLI options

- **Production server**: `npx ts-node src/server.ts`

### Building and Testing

- **TypeScript compilation**: `npm run build`
  - Compiles TypeScript to `lib/` directory
  - Uses `tsconfig.json` (excludes Angular webapp)

- **Run tests**: `npm test`
  - Uses Jest with `ts-jest` preset
  - Test files: `tests/**/*.test.ts`
  - Generates coverage report

- **Angular development**: Not yet fully configured
  - Angular files are in `src/webapp/`
  - Configuration: `angular.json`
  - The Angular app is being migrated from an older implementation

### Docker

- **Build image**: `npm run docker:build`
- **Run container**: `npm run docker:run`
  - Mounts three volumes: `/input` (template), `/output` (sermon folders), `/songs`
  - See `package.json` scripts for full docker commands with volume mappings
- **Build and run**: `npm run docker:build-n-run`
- **Stop and restart**: `npm run docker:stop-n-run`

Note: The Docker image runs with ts-node (not compiled), so dev dependencies are included.

## Architecture

### Dual Application Structure

This project has two distinct parts that are loosely coupled:

1. **Node.js Backend** (`src/server.ts`, `src/sunday-powerpoints.ts`, `src/index.ts`)
   - Express server serving API endpoints and static files
   - Socket.IO for real-time updates to the frontend
   - Automated folder/file management for Sunday PowerPoint presentations
   - Windows shortcut (.lnk) management via `windows-shortcuts` library
   - Scheduled jobs using `node-schedule`

2. **Angular Frontend** (`src/webapp/`)
   - Material Design UI (`@angular/material`)
   - Components for song selection, file browsing, theme toggle
   - Socket service for real-time server communication
   - Currently in migration (see commit `57bb0d8`)

### Core Backend Classes (server.ts)

- **`FolderCache`**: Caches `SundayFolder` objects by date (YYYYMMDD format)
  - Has a `garbage` collection for recently deleted items (10-minute retention)

- **`SundayFolder`**: Represents a Sunday's sermon folder
  - Properties: `name`, `path`, `presentation`, `thumbnail`, `video[]`, `song1-3`, `choruses[]`, `hasNotes`
  - Scans folders to detect PowerPoint files, videos, song shortcuts, notes

- **`Video`**: Represents video file metadata

- **`Song`**: Represents a song shortcut with metadata (name, target path, CCLI)

- **`Chorus`**: Represents a chorus shortcut

### Key Backend Utilities (sunday-powerpoints.ts)

- **`sundaysInMonth(m, y)`**: Returns array of Sunday day numbers for a given month/year
  - Handles DST edge cases (see commit `b7612bf`)

- **`createShortcut(filename, description, targetFilePath, rootPath?)`**: Creates Windows .lnk files
  - Replaces absolute OneDrive paths with `%OneDriveConsumer%` environment variable

- **`fixExistingShortcuts(directory, rootPath?)`**: Batch updates shortcuts to use environment variables

- **`resolveToAbsolutePath(path)`**: Expands Windows environment variables in paths

### Angular Components (src/webapp/)

- **Admin Component**: Administrative interface (structure minimal)
- **Song Library**: Browse/search available songs
- **Song Select Modal**: UI for selecting songs for a specific Sunday
- **Song Badge**: Display component for song metadata
- **Upcoming Songs**: Shows songs scheduled for upcoming Sundays
- **Output Folders Sidebar**: Browse sermon output folders
- **Notes Modal**: Edit/view notes for a Sunday
- **Theme Toggle**: Dark/light mode switcher

### Environment Configuration

The application uses `.env` for configuration (excluded from Docker image):

```
PORT=8080
TEMPLATE_FILE=Sunday Template.pptx
EXT=pptx
TEMPLATE_DIRECTORY=<path to template directory>
OUTPUT_DIRECTORY=<path to output directory>
SONGS_DIRECTORY=<path to songs directory>
ROOT_PATH=%OneDriveConsumer%
```

Docker runtime overrides these via `-e` flags and volume mounts.

### TypeScript Configuration

- **Root `tsconfig.json`**: For Node.js backend (outputs to `lib/`)
- **`src/webapp/tsconfig.json`**: Angular-specific configuration
- **`src/webapp/tsconfig.app.json`**: Angular build config
- **`src/webapp/tsconfig.spec.json`**: Angular test config

The root TypeScript config excludes the Angular webapp to avoid conflicts.

## Important Patterns

### Date Handling

- Dates are stored in `YYYYMMDD` format for folder names
- Use `moment` for date calculations (see `src/index.ts`)
- The `sundaysInMonth()` function accounts for DST transitions

### Windows-Specific Code

- This application is **Windows-dependent** due to `windows-shortcuts` library
- Shortcuts use Windows environment variable syntax (`%OneDriveConsumer%`)
- Docker feasibility depends on whether shortcuts work in Linux containers (see roadmap)

### File System Operations

- Template files are located via `TEMPLATE_DIRECTORY + TEMPLATE_FILE`
- Output folders follow pattern: `OUTPUT_DIRECTORY/YYYYMMDD/`
- Songs are referenced by creating shortcuts (.lnk files) in each Sunday folder

### Socket.IO Communication

The backend emits real-time updates to connected clients:
- Folder updates when new folders are created
- Song selection changes
- File system changes

## Testing Strategy

Tests are located in `tests/` and use Jest with TypeScript support. The test configuration (`jest.config.js`) runs tests matching `**/tests/**/*.test.ts`.

Run a single test: `npx jest tests/sunday-powerpoints.test.ts`

## Roadmap and Active Tasks

**See `src/roadmap.md` for complete feature planning, implementation details, and current status.**

When starting work, create a todo list from current GitHub issues using: `gh issue list --repo pcnate/sunday-powerpoints`
