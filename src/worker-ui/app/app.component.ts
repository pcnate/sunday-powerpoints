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
import { Subject, takeUntil, interval } from 'rxjs';


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
  shift_start: string;
  shift_end: string;
}


/**
 * Worker config from GET /api/config.
 */
interface WorkerConfig {
  server: { url: string };
  worker: { id: string; types: string[] };
  schedule: { shift_start: string; shift_end: string };
  timing: {
    poll_interval_secs: number;
    heartbeat_interval_secs: number;
    job_delay_secs: number;
  };
  tools: {
    whisper_path: string;
    whisper_model: string;
    claude_path: string;
    ffmpeg_path: string;
  };
  web_ui: { port: number };
}


/**
 * Root component for the worker config UI.
 * Displays real-time status and an editable configuration form.
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
  ],
  templateUrl: './app.component.html',
  styleUrls: [ './app.component.scss' ]
})
export class AppComponent implements OnInit, OnDestroy {

  status: WorkerStatus | null = null;
  connected = false;

  // Config form
  serverUrl = '';
  workerId = '';
  typeAlignment = false;
  typeTranscription = false;
  typeClaude = false;
  shiftStart = '';
  shiftEnd = '';
  pollInterval = 30;
  heartbeatInterval = 30;
  jobDelay = 10;
  whisperPath = '';
  whisperModel = '';
  claudePath = '';
  ffmpegPath = '';

  saving = false;

  private destroy$ = new Subject<void>();


  constructor(
    private http: HttpClient,
    private snackBar: MatSnackBar
  ) {}


  /**
   * Initialize status polling and load config.
   */
  ngOnInit(): void {
    this.loadStatus();
    this.loadConfig();

    interval( 3000 ).pipe( takeUntil( this.destroy$ ) )
      .subscribe( () => this.loadStatus() );
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
          this.typeAlignment = c.worker.types.includes( 'video-alignment' );
          this.typeTranscription = c.worker.types.includes( 'transcription' );
          this.typeClaude = c.worker.types.includes( 'claude-processing' );
          this.shiftStart = c.schedule.shift_start;
          this.shiftEnd = c.schedule.shift_end;
          this.pollInterval = c.timing.poll_interval_secs;
          this.heartbeatInterval = c.timing.heartbeat_interval_secs;
          this.jobDelay = c.timing.job_delay_secs;
          this.whisperPath = c.tools.whisper_path;
          this.whisperModel = c.tools.whisper_model;
          this.claudePath = c.tools.claude_path;
          this.ffmpegPath = c.tools.ffmpeg_path;
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
    if ( this.typeAlignment ) types.push( 'video-alignment' );
    if ( this.typeTranscription ) types.push( 'transcription' );
    if ( this.typeClaude ) types.push( 'claude-processing' );

    const body = {
      server_url: this.serverUrl,
      worker_id: this.workerId,
      types,
      shift_start: this.shiftStart,
      shift_end: this.shiftEnd,
      poll_interval_secs: this.pollInterval,
      heartbeat_interval_secs: this.heartbeatInterval,
      job_delay_secs: this.jobDelay,
      whisper_path: this.whisperPath,
      whisper_model: this.whisperModel,
      claude_path: this.claudePath,
      ffmpeg_path: this.ffmpegPath,
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
