# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Sunday-powerpoints is a hybrid Node.js/Angular application designed to automate PowerPoint management for weekly church services. The system creates folders for each Sunday of the month, manages song shortcuts, tracks presentation files, and provides a web interface for song selection and file management. It includes a job queue system for automated video processing (alignment, transcription, and AI-powered metadata extraction).

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

- **Run single test**: `npx jest tests/jobs.test.ts`

- **Angular development**: Not yet fully configured
  - Angular files are in `src/webapp/`
  - Configuration: `angular.json`
  - The Angular app is being migrated from an older implementation

### Rust Worker

- **Build (debug)**: `cargo check --manifest-path worker/Cargo.toml`
- **Build (release)**: `cargo build --release --manifest-path worker/Cargo.toml`
- **Run**: `cargo run --manifest-path worker/Cargo.toml`
- **Lint**: `cargo clippy --manifest-path worker/Cargo.toml -- -W clippy::all`
- **Release binary**: `worker/target/release/sunday-worker.exe` (~6 MB)
- **Config file**: `%APPDATA%\sunday-worker\config.toml` (auto-created on first run)
- **Web UI**: `http://localhost:9090` (configurable)

### API Documentation

- **Swagger UI**: Available at `http://localhost:8080/api-docs` when server is running
- **OpenAPI spec**: `src/openapi.yaml` (also served at `/openapi.yaml`)

## Architecture

### Three-Part Application Structure

1. **Node.js Backend** (`src/server.ts`, `src/sunday-powerpoints.ts`, `src/index.ts`)
   - Express server serving API endpoints and static files
   - Socket.IO for real-time updates to the frontend
   - MySQL-backed job queue for video processing pipeline
   - Automated folder/file management for Sunday PowerPoint presentations
   - Windows shortcut (.lnk) management via `windows-shortcuts` library
   - Scheduled jobs using `node-schedule`
   - Swagger UI at `/api-docs`

2. **Angular Frontend** (`src/webapp/`)
   - Material Design UI (`@angular/material`)
   - Components for song selection, file browsing, theme toggle
   - Socket service for real-time server communication
   - Currently in migration (see commit `57bb0d8`)

3. **Rust Worker** (`worker/`)
   - Windows system tray application for GPU-equipped machine
   - Processes transcription and AI jobs during configured shift windows
   - Communicates only via Express server REST API (no direct DB access)
   - Two-thread architecture: main thread (winit + tray-icon), tokio thread (scheduler + HTTP)
   - Config stored at `%APPDATA%\sunday-worker\config.toml`
   - Embedded web UI at `http://localhost:9090` for config management

### Job Queue System

The job queue processes a video pipeline for sermon recordings:

**Pipeline**: Raw MP4s → Kdenlive project → Manual editing → Production MP4 → Whisper transcription → Claude Code processing

**Key files:**
- `src/db.ts` — MySQL connection pool with automatic migration runner
- `src/types/job.ts` — TypeScript interfaces for all job-related types
- `src/routes/jobs.ts` — Express Router for job CRUD + worker endpoints
- `src/migrations/*.sql` — Database schema migrations (run automatically on startup)
- `src/openapi.yaml` — Full API documentation

**Job types:**
- `video-alignment` — FFmpeg scene detection + Kdenlive project generation
- `transcription` — Whisper speech-to-text on production MP4
- `claude-processing` — Claude Code extracts sermon metadata from VTT

**Job lifecycle:** `pending` → `processing` (claimed by worker) → `completed`/`failed`

**Auto-detection:** `scanFolders()` runs every minute and:
- Creates `video-alignment` jobs when 2+ raw MP4s found in `Vids/` without a production file
- Creates `transcription` jobs when `YYYYMMDD-production.mp4` appears

**Job chaining:** Completing a `transcription` job auto-creates a `claude-processing` job.

**Stale recovery:** Every 2 minutes, jobs with heartbeats older than 2 minutes are recovered (retried or failed).

### Rust Worker Architecture (`worker/`)

**Source files:**
- `src/main.rs` — Entry point: config loading, thread spawning, tray + scheduler wiring
- `src/config.rs` — TOML config deserialization, load/save, `%APPDATA%` path resolution
- `src/state.rs` — `AppState` (shared via `Arc<RwLock>`), `SchedulerPhase` enum, `TrayCommand` enum
- `src/api_client.rs` — reqwest HTTP client wrapping all Express job API calls
- `src/scheduler.rs` — Shift window logic + polling state machine
- `src/heartbeat.rs` — Background heartbeat task with `CancellationToken`
- `src/job_runner.rs` — Job execution orchestrator (dispatches to runners)
- `src/tray.rs` — System tray icon + menu using winit + tray-icon + muda
- `src/web_ui.rs` — Embedded axum config web server on port 9090
- `src/runners/{alignment,transcription,claude}.rs` — **STUB** runners (sleep 5s, return success)

**Scheduler state machine:**
`OFF_SHIFT → IDLE → POLLING → EXECUTING → COOLDOWN → IDLE` with `PAUSED` toggle

**Communication channels:**
- `tokio::sync::mpsc` — tray UI commands → scheduler (pause, quit, open config, reload)
- `Arc<AppState>` with `tokio::sync::RwLock` — scheduler → tray (phase, job info, stats)
- `CancellationToken` — graceful shutdown propagation

**Tray icon colors:** Green (active/idle), Yellow (executing), Gray (off-shift/paused), Red (disconnected)

### Database

- **MySQL** on NAS, connection configured via `MYSQL_*` env vars
- **Migrations** run automatically on server startup from `src/migrations/`
- Migration files are numbered sequentially (e.g., `001_create_jobs_tables.sql`)
- Each migration runs once; tracking via `migrations` table
- Add new migrations as new numbered `.sql` files — never modify applied ones
- Server continues without job system if MySQL is unavailable

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

The application uses `.env` for configuration (see `.env.example` for all options):

```
# Server
PORT=8080
TEMPLATE_FILE=Sunday Template.pptx
EXT=pptx
TEMPLATE_DIRECTORY=<path to template directory>
OUTPUT_DIRECTORY=<path to output directory>
SONGS_DIRECTORY=<path to songs directory>
ROOT_PATH=%OneDriveConsumer%

# MySQL (optional - server runs without it)
MYSQL_HOST=192.168.1.50
MYSQL_PORT=3306
MYSQL_USER=sunday_powerpoints
MYSQL_PASSWORD=<password>
MYSQL_DATABASE=sunday_powerpoints

# Job system tuning
HEARTBEAT_TIMEOUT_MS=120000
```

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
- Docker was abandoned due to Windows shortcuts incompatibility

### File System Operations

- Template files are located via `TEMPLATE_DIRECTORY + TEMPLATE_FILE`
- Output folders follow pattern: `OUTPUT_DIRECTORY/YYYYMMDD/`
- Songs are referenced by creating shortcuts (.lnk files) in each Sunday folder
- Video files live in `OUTPUT_DIRECTORY/YYYYMMDD/Vids/`
- Production MP4 naming: `YYYYMMDD-production.mp4`
- Kdenlive projects: `YYYYMMDD.kdenlive`

### Socket.IO Communication

The backend emits real-time updates to connected clients:
- `folder-changes` — Folder updates when content changes
- `job:created` — New job added to queue
- `job:updated` — Job status changed (claimed, completed, failed, recovered)
- `job:log` — New log entries for a job

### API Routes

All REST endpoints documented in Swagger UI (`/api-docs`). Key groups:
- `/api/folders` — Folder listing and thumbnails
- `/api/songs` — Song library and selection
- `/api/notes` — Sermon notes CRUD
- `/api/jobs` — Job queue management
- `/api/jobs/next`, `/api/jobs/:id/heartbeat`, etc. — Worker communication endpoints

## Testing Strategy

Tests are located in `tests/` and use Jest with TypeScript support. The test configuration (`jest.config.js`) runs tests matching `**/tests/**/*.test.ts`.

- `tests/sunday-powerpoints.test.ts` — Core utility tests
- `tests/jobs.test.ts` — Job queue system tests (DB mocked, logic verified)

Run all tests: `npm test`
Run a single test: `npx jest tests/jobs.test.ts`

## Roadmap and Active Tasks

**See `src/roadmap.md` for complete feature planning, implementation details, and current status.**

When starting work, create a todo list from current GitHub issues using: `gh issue list --repo pcnate/sunday-powerpoints`
