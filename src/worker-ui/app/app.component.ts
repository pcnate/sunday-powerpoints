import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MatCardModule } from '@angular/material/card';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { MatSnackBar, MatSnackBarModule } from '@angular/material/snack-bar';
import { MatDividerModule } from '@angular/material/divider';
import { MatTooltipModule } from '@angular/material/tooltip';
import { MatAutocompleteModule } from '@angular/material/autocomplete';
import { MatTabsModule } from '@angular/material/tabs';
import { MatSlideToggleModule } from '@angular/material/slide-toggle';
import { MatSelectModule } from '@angular/material/select';
import { MatDialog, MatDialogModule } from '@angular/material/dialog';
import { Subject, takeUntil, interval } from 'rxjs';
import { FileBrowserDialogComponent, FileBrowserDialogData } from './file-browser-dialog.component';


/**
 * Worker status from GET /api/status.
 */
interface WorkerStatus {
  phase: string;
  connected: boolean;
  current_job_id: number | null;
  current_job_type: string | null;
  jobs_completed: number;
  jobs_failed: number;
  uptime_secs: number;
  worker_id: string;
  server_url: string;
  schedule_enabled: boolean;
  shift_start: string;
  shift_end: string;
  force_on_shift: boolean;
}


/**
 * Worker config from GET /api/config.
 */
interface WorkerConfig {
  server: { url: string };
  worker: { id: string; types: string[] };
  schedule: { enabled: boolean; shift_start: string; shift_end: string };
  timing: {
    poll_interval_secs: number;
    heartbeat_interval_secs: number;
    job_delay_secs: number;
  };
  tools: {
    whisper_path: string;
    whisper_model: string;
    whisper_compute_type?: string;
    claude_path: string;
    ffmpeg_path: string;
    ffprobe_path: string;
    melt_path: string;
    melt_video_bitrate: string;
    melt_audio_bitrate: string;
    tool_env?: Record<string, Record<string, string>>;
  };
  paths: { output_directory: string };
  web_ui: { port: number };
}


/**
 * Job record from GET /api/queue.
 */
interface QueueJob {
  id: number;
  type: string;
  sunday_date: string;
  status: string;
  input_path: string;
  output_path?: string;
  worker_id?: string;
  retry_count: number;
  max_retries: number;
  error_message?: string;
  created_at?: string;
  started_at?: string;
  completed_at?: string;
}


/**
 * Log entry from GET /api/queue/:id.
 */
interface JobLog {
  level: string;
  message: string;
  created_at?: string;
}


/**
 * Result of a tool test.
 */
interface TestResult {
  success: boolean;
  output: string;
}


/**
 * Root component for the worker config UI.
 * Displays real-time status and a tabbed configuration form.
 */
@Component({
  selector: 'app-root',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatCardModule,
    MatButtonModule,
    MatIconModule,
    MatFormFieldModule,
    MatInputModule,
    MatCheckboxModule,
    MatSnackBarModule,
    MatDividerModule,
    MatTooltipModule,
    MatTabsModule,
    MatSlideToggleModule,
    MatSelectModule,
    MatDialogModule,
    MatAutocompleteModule,
  ],
  templateUrl: './app.component.html',
  styleUrls: [ './app.component.scss' ]
})
export class AppComponent implements OnInit, OnDestroy {

  status: WorkerStatus | null = null;
  connected = false;

  // Config form — Settings tab
  serverUrl = '';
  workerId = '';
  typeTranscode = false;
  typeTranscription = false;
  typeClaude = false;
  typeFfprobe = false;
  scheduleEnabled = true;
  shiftStart = '';
  shiftEnd = '';
  pollInterval = 30;
  heartbeatInterval = 30;
  jobDelay = 10;

  // Config form — Tool tabs
  meltPath = '';
  meltVideoBitrate = '5000k';
  meltAudioBitrate = '192k';
  whisperPath = '';
  whisperModel = '';
  whisperComputeType = 'float16';
  claudePath = '';
  ffmpegPath = '';
  ffprobePath = '';
  outputDirectory = '';

  // Per-tool environment variables (KEY=VALUE text per tool)
  envWhisper = '';
  envClaude = '';
  envFfprobe = '';
  envFfmpeg = '';
  envMelt = '';

  whisperModels = [
    'large-v3', 'large-v3-turbo', 'turbo', 'large-v2', 'large-v1',
    'distil-large-v3.5', 'distil-large-v3', 'distil-large-v2',
    'medium', 'medium.en', 'small', 'small.en',
    'distil-medium.en', 'distil-small.en',
    'base', 'base.en', 'tiny', 'tiny.en',
  ];

  whisperComputeTypes = [
    'float16', 'float32', 'int8', 'int8_float16', 'int8_float32', 'auto', 'default',
  ];

  saving = false;
  selectedTabIndex = 0;

  // Job Queue tab
  jobs: QueueJob[] = [];
  queueTotal = 0;
  queueFilter = 'pending,processing';
  queueTypeFilter = '';
  selectedJob: QueueJob | null = null;
  jobLogs: JobLog[] = [];
  loadingDetail = false;

  // Tool testing
  testResults: Record<string, TestResult | null> = {};
  testing: Record<string, boolean> = {};

  private destroy$ = new Subject<void>();


  constructor(
    private http: HttpClient,
    private snackBar: MatSnackBar,
    private dialog: MatDialog,
  ) {}


  /**
   * Initialize status polling and load config.
   */
  ngOnInit(): void {
    this.loadStatus();
    this.loadConfig();
    this.loadQueue();

    interval( 3000 ).pipe( takeUntil( this.destroy$ ) )
      .subscribe( () => this.loadStatus() );

    interval( 10000 ).pipe( takeUntil( this.destroy$ ) )
      .subscribe( () => {
        if ( this.selectedTabIndex === 0 ) {
          this.loadQueue();
          if ( this.selectedJob ) {
            this.refreshJobDetail( this.selectedJob.id );
          }
        }
      } );
  }


  /**
   * Clean up subscriptions.
   */
  ngOnDestroy(): void {
    this.destroy$.next();
    this.destroy$.complete();
  }


  /**
   * Fetch worker status from the local Rust worker API.
   */
  loadStatus(): void {
    this.http.get<WorkerStatus>( '/api/status' )
      .subscribe({
        next: ( s ) => {
          this.status = s;
          this.connected = true;
        },
        error: () => { this.connected = false; },
      });
  }


  /**
   * Fetch worker config and populate form fields.
   */
  loadConfig(): void {
    this.http.get<WorkerConfig>( '/api/config' )
      .subscribe({
        next: ( c ) => {
          this.serverUrl = c.server.url;
          this.workerId = c.worker.id;
          this.typeTranscode = c.worker.types.includes( 'transcode' );
          this.typeTranscription = c.worker.types.includes( 'transcription' );
          this.typeClaude = c.worker.types.includes( 'claude-processing' );
          this.typeFfprobe = c.worker.types.includes( 'ffprobe' );
          this.scheduleEnabled = c.schedule.enabled !== false;
          this.shiftStart = c.schedule.shift_start;
          this.shiftEnd = c.schedule.shift_end;
          this.pollInterval = c.timing.poll_interval_secs;
          this.heartbeatInterval = c.timing.heartbeat_interval_secs;
          this.jobDelay = c.timing.job_delay_secs;
          this.whisperPath = c.tools.whisper_path;
          this.whisperModel = c.tools.whisper_model;
          this.whisperComputeType = c.tools.whisper_compute_type || 'float16';
          this.claudePath = c.tools.claude_path;
          this.ffmpegPath = c.tools.ffmpeg_path;
          this.ffprobePath = c.tools.ffprobe_path || '';
          this.outputDirectory = c.paths?.output_directory || '';
          this.meltPath = c.tools.melt_path;
          this.meltVideoBitrate = c.tools.melt_video_bitrate || '5000k';
          this.meltAudioBitrate = c.tools.melt_audio_bitrate || '192k';

          // Load per-tool env vars
          const toolEnv = c.tools.tool_env || {};
          this.envWhisper = this.envToText( toolEnv[ 'whisper' ] || {} );
          this.envClaude = this.envToText( toolEnv[ 'claude' ] || {} );
          this.envFfprobe = this.envToText( toolEnv[ 'ffprobe' ] || {} );
          this.envFfmpeg = this.envToText( toolEnv[ 'ffmpeg' ] || {} );
          this.envMelt = this.envToText( toolEnv[ 'melt' ] || {} );
        },
        error: () => {
          this.snackBar.open( 'Could not load worker config', 'Dismiss', { duration: 4000 } );
        },
      });
  }


  /**
   * Save config form to the worker.
   */
  saveConfig(): void {
    const types: string[] = [];
    if ( this.typeTranscode ) types.push( 'transcode' );
    if ( this.typeTranscription ) types.push( 'transcription' );
    if ( this.typeClaude ) types.push( 'claude-processing' );
    if ( this.typeFfprobe ) types.push( 'ffprobe' );

    const body = {
      server_url: this.serverUrl,
      worker_id: this.workerId,
      types,
      schedule_enabled: this.scheduleEnabled,
      shift_start: this.shiftStart,
      shift_end: this.shiftEnd,
      poll_interval_secs: this.pollInterval,
      heartbeat_interval_secs: this.heartbeatInterval,
      job_delay_secs: this.jobDelay,
      whisper_path: this.whisperPath,
      whisper_model: this.whisperModel,
      whisper_compute_type: this.whisperComputeType,
      claude_path: this.claudePath,
      ffmpeg_path: this.ffmpegPath,
      ffprobe_path: this.ffprobePath,
      output_directory: this.outputDirectory,
      melt_path: this.meltPath,
      melt_video_bitrate: this.meltVideoBitrate,
      melt_audio_bitrate: this.meltAudioBitrate,
      tool_env: this.buildToolEnv(),
    };

    this.saving = true;
    this.http.post<{ success: boolean }>( '/api/config', body )
      .subscribe({
        next: () => {
          this.saving = false;
          this.snackBar.open( 'Configuration saved', '', { duration: 3000 } );
        },
        error: ( err ) => {
          this.saving = false;
          this.snackBar.open(
            'Failed to save: ' + ( err.error?.error || 'Unknown error' ),
            'Dismiss',
            { duration: 5000 }
          );
        },
      });
  }


  /**
   * Toggle the scheduler's paused state.
   */
  togglePause(): void {
    this.http.post( '/api/scheduler/pause', {} )
      .subscribe({
        next: () => setTimeout( () => this.loadStatus(), 300 ),
        error: () => this.snackBar.open( 'Failed to toggle pause', 'Dismiss', { duration: 3000 } ),
      });
  }


  /**
   * Toggle force on-shift override.
   */
  toggleForceOnShift(): void {
    this.http.post( '/api/scheduler/force-on-shift', {} )
      .subscribe({
        next: () => setTimeout( () => this.loadStatus(), 300 ),
        error: () => this.snackBar.open( 'Failed to toggle force on-shift', 'Dismiss', { duration: 3000 } ),
      });
  }


  /**
   * Open a file browser dialog and set the selected path on a field.
   *
   * @param field - property name to set the selected path on
   * @param title - dialog title
   */
  browse( field: string, title: string ): void {
    const currentValue = ( this as Record<string, unknown> )[ field ] as string || '';

    const data: FileBrowserDialogData = {
      title,
      startPath: currentValue,
      extensions: [ 'exe' ],
    };

    const ref = this.dialog.open( FileBrowserDialogComponent, {
      width: '600px',
      data,
    } );

    ref.afterClosed().subscribe( ( result: string | null ) => {
      if ( result ) {
        ( this as Record<string, unknown> )[ field ] = result;
      }
    } );
  }


  /**
   * Open a directory browser dialog and set the selected path on a field.
   *
   * @param field - property name to set the selected path on
   * @param title - dialog title
   */
  browseDirectory( field: string, title: string ): void {
    const currentValue = ( this as Record<string, unknown> )[ field ] as string || '';

    const data: FileBrowserDialogData = {
      title,
      startPath: currentValue,
      directoryMode: true,
    };

    const ref = this.dialog.open( FileBrowserDialogComponent, {
      width: '600px',
      data,
    } );

    ref.afterClosed().subscribe( ( result: string | null ) => {
      if ( result ) {
        ( this as Record<string, unknown> )[ field ] = result;
      }
    } );
  }


  /**
   * Test a tool by running it with a version/help flag.
   *
   * @param jobType - which tool to test (transcode, transcription, claude-processing, ffmpeg)
   */
  testTool( jobType: string ): void {
    let command = '';
    let args: string[] = [];

    switch ( jobType ) {
      case 'transcode':
        command = this.meltPath;
        args = [ '-version' ];
        break;
      case 'ffmpeg':
        command = this.ffmpegPath;
        args = [ '-version' ];
        break;
      case 'transcription':
        command = this.whisperPath;
        args = [ '--help' ];
        break;
      case 'claude-processing':
        command = this.claudePath;
        args = [ '--version' ];
        break;
      case 'ffprobe':
        command = this.ffprobePath;
        args = [ '-version' ];
        break;
      default:
        return;
    }

    // Resolve env vars for the tool being tested
    const envTextMap: Record<string, string> = {
      transcode: this.envMelt,
      transcription: this.envWhisper,
      'claude-processing': this.envClaude,
      ffprobe: this.envFfprobe,
      ffmpeg: this.envFfmpeg,
    };
    const env = this.parseEnvText( envTextMap[ jobType ] || '' );

    this.testing[ jobType ] = true;
    this.testResults[ jobType ] = null;

    this.http.post<{ success: boolean; output?: string; error?: string }>( '/api/test-tool', {
      command,
      args,
      env: Object.keys( env ).length > 0 ? env : undefined,
    } ).subscribe({
      next: ( res ) => {
        this.testing[ jobType ] = false;
        this.testResults[ jobType ] = {
          success: res.success,
          output: res.output || res.error || '',
        };
      },
      error: ( err ) => {
        this.testing[ jobType ] = false;
        this.testResults[ jobType ] = {
          success: false,
          output: err.error?.error || 'Request failed',
        };
      },
    });
  }


  /**
   * Request the scheduler to run exactly one job cycle.
   */
  runOnce(): void {
    this.http.post( '/api/scheduler/run-once', {} )
      .subscribe({
        next: () => {
          this.snackBar.open( 'Run-once requested', '', { duration: 3000 } );
          setTimeout( () => this.loadStatus(), 1000 );
        },
        error: () => this.snackBar.open( 'Run-once failed', 'Dismiss', { duration: 3000 } ),
      });
  }


  /**
   * Parse a KEY=VALUE text block into a map.
   *
   * @param text - newline-separated KEY=VALUE pairs
   * @returns parsed map
   */
  parseEnvText( text: string ): Record<string, string> {
    const env: Record<string, string> = {};
    for ( const line of text.split( '\n' ) ) {
      const trimmed = line.trim();
      if ( !trimmed || !trimmed.includes( '=' ) ) continue;
      const idx = trimmed.indexOf( '=' );
      const key = trimmed.substring( 0, idx ).trim();
      const value = trimmed.substring( idx + 1 ).trim();
      if ( key ) env[ key ] = value;
    }
    return env;
  }


  /**
   * Convert an env var map to KEY=VALUE text.
   *
   * @param env - the env var map
   * @returns newline-separated KEY=VALUE string
   */
  envToText( env: Record<string, string> ): string {
    return Object.entries( env ).map( ([ k, v ]) => `${ k }=${ v }` ).join( '\n' );
  }


  /**
   * Build the tool_env object from all tool env text fields.
   *
   * @returns per-tool env var maps
   */
  buildToolEnv(): Record<string, Record<string, string>> {
    const result: Record<string, Record<string, string>> = {};
    const fields: [ string, string ][] = [
      [ 'whisper', this.envWhisper ],
      [ 'claude', this.envClaude ],
      [ 'ffprobe', this.envFfprobe ],
      [ 'ffmpeg', this.envFfmpeg ],
      [ 'melt', this.envMelt ],
    ];
    for ( const [ tool, text ] of fields ) {
      const env = this.parseEnvText( text );
      if ( Object.keys( env ).length > 0 ) {
        result[ tool ] = env;
      }
    }
    return result;
  }


  /**
   * Fetch the job queue list from the worker proxy.
   */
  loadQueue(): void {
    let url = '/api/queue?limit=50';
    if ( this.queueFilter ) url += `&status=${ this.queueFilter }`;
    if ( this.queueTypeFilter ) url += `&type=${ this.queueTypeFilter }`;

    this.http.get<{ jobs: QueueJob[]; total: number }>( url )
      .subscribe({
        next: ( data ) => {
          this.jobs = data.jobs || [];
          this.queueTotal = data.total || 0;
        },
        error: () => {
          this.jobs = [];
          this.queueTotal = 0;
        },
      });
  }


  /**
   * Show detail panel for a job (fetch job + logs).
   *
   * @param id - job ID to fetch
   */
  showJobDetail( id: number ): void {
    this.loadingDetail = true;
    this.http.get<{ job: QueueJob; logs: JobLog[] }>( `/api/queue/${ id }` )
      .subscribe({
        next: ( data ) => {
          this.selectedJob = data.job || data as unknown as QueueJob;
          this.jobLogs = data.logs || [];
          this.loadingDetail = false;
        },
        error: () => {
          this.loadingDetail = false;
        },
      });
  }


  /**
   * Silently refresh the detail panel without resetting loadingDetail.
   *
   * @param id - job ID to refresh
   */
  private refreshJobDetail( id: number ): void {
    this.http.get<{ job: QueueJob; logs: JobLog[] }>( `/api/queue/${ id }` )
      .subscribe({
        next: ( data ) => {
          this.selectedJob = data.job || data as unknown as QueueJob;
          this.jobLogs = data.logs || [];
        },
      });
  }


  /**
   * Close the job detail panel.
   */
  closeJobDetail(): void {
    this.selectedJob = null;
    this.jobLogs = [];
  }


  /**
   * Cancel a job by ID.
   *
   * @param id - job ID to cancel
   */
  cancelJob( id: number ): void {
    this.http.put( `/api/queue/${ id }/cancel`, {} )
      .subscribe({
        next: () => {
          this.snackBar.open( `Job #${ id } cancelled`, '', { duration: 3000 } );
          this.loadQueue();
          if ( this.selectedJob?.id === id ) {
            this.showJobDetail( id );
          }
        },
        error: () => {
          this.snackBar.open( `Failed to cancel job #${ id }`, 'Dismiss', { duration: 4000 } );
        },
      });
  }


  /**
   * Check if a log message looks like a CLI command.
   *
   * @param msg - log message to check
   * @returns true if the message looks like a command invocation
   */
  isCommandLine( msg: string ): boolean {
    return /^[A-Za-z]:[\\\/]/.test( msg );
  }


  /**
   * Get CSS class for the status indicator dot.
   *
   * @returns CSS class name
   */
  statusDotClass(): string {
    if ( !this.connected || !this.status ) return 'dot-disconnected';
    if ( !this.status.connected ) return 'dot-disconnected';

    switch ( this.status.phase ) {
      case 'Executing': return 'dot-executing';
      case 'Idle':
      case 'Polling':
      case 'Cooldown': return 'dot-active';
      case 'Paused': return 'dot-paused';
      default: return 'dot-off';
    }
  }


  /**
   * Format uptime seconds as a readable string.
   *
   * @param secs - uptime in seconds
   * @returns formatted string like "3h 12m"
   */
  formatUptime( secs: number ): string {
    const hrs = Math.floor( secs / 3600 );
    const mins = Math.floor( ( secs % 3600 ) / 60 );
    return `${ hrs }h ${ mins }m`;
  }
}
