import { Component, Inject, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MatDialogModule, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatSelectModule } from '@angular/material/select';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { DraggableDialogDirective } from './draggable-dialog.directive';


/**
 * Data passed to the queue job dialog.
 */
export interface QueueJobDialogData {
  folderName: string;
  jobType: 'transcription' | 'claude-processing';
}


/**
 * Previous job record from the history endpoint.
 */
interface PreviousJob {
  id: number;
  type: string;
  status: string;
  error_message: string | null;
  worker_id: string | null;
  input_path: string;
  output_path: string | null;
  created_at: string;
  started_at: string | null;
  completed_at: string | null;
  retry_count: number;
  max_retries: number;
}


/**
 * Video file info from the video-files endpoint.
 */
interface VideoFile {
  filename: string;
  size: number;
  type: string;
}


/**
 * Dialog for queuing transcription or claude-processing jobs.
 *
 * Shows previous job history if available. For transcription jobs,
 * lets the user select which video file to transcribe.
 */
@Component({
  selector: 'app-queue-job-dialog',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatDialogModule,
    MatButtonModule,
    MatIconModule,
    MatSelectModule,
    MatFormFieldModule,
    MatProgressSpinnerModule,
    DraggableDialogDirective,
  ],
  template: `
    <h2 mat-dialog-title appDraggableDialog>
      <mat-icon>{{ icon }}</mat-icon>
      {{ title }}
    </h2>

    <mat-dialog-content>
      <div *ngIf="loading" class="loading">
        <mat-spinner diameter="30"></mat-spinner>
      </div>

      <ng-container *ngIf="!loading">
        <p class="folder-label">{{ data.folderName }}</p>

        <!-- Video selector for transcription -->
        <ng-container *ngIf="data.jobType === 'transcription' && videoFiles.length > 0">
          <mat-form-field appearance="outline" class="full-width">
            <mat-label>Video to transcribe</mat-label>
            <mat-select [(ngModel)]="selectedVideo">
              <mat-option *ngFor="let f of videoFiles" [value]="f.filename">
                {{ f.filename }} ({{ formatSize( f.size ) }})
              </mat-option>
            </mat-select>
          </mat-form-field>
        </ng-container>

        <div *ngIf="data.jobType === 'transcription' && videoFiles.length === 0 && !error" class="empty-state">
          No video files found in Vids/ directory.
        </div>

        <!-- Previous job history -->
        <div *ngIf="previousJobs.length > 0" class="history-section">
          <h3>Previous Jobs</h3>
          <div class="history-list">
            <div *ngFor="let job of previousJobs" class="history-item">
              <div class="history-header">
                <span class="badge" [ngClass]="'badge-' + job.status">#{{ job.id }} — {{ job.status }}</span>
                <span class="history-date">{{ formatDate( job.created_at ) }}</span>
              </div>
              <div *ngIf="job.input_path" class="history-path">
                <span class="path-label">Input:</span> {{ job.input_path }}
              </div>
              <div *ngIf="job.worker_id" class="history-detail">
                Worker: {{ job.worker_id }}
              </div>
              <div *ngIf="job.error_message" class="history-error">
                {{ job.error_message }}
              </div>
            </div>
          </div>
        </div>

        <div *ngIf="error" class="status-message error">{{ error }}</div>
        <div *ngIf="successMessage" class="status-message success">{{ successMessage }}</div>
      </ng-container>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button (click)="dialogRef.close( queued )">
        {{ queued ? 'Close' : 'Cancel' }}
      </button>
      <button mat-raised-button color="primary"
              *ngIf="!queued"
              [disabled]="!canQueue || submitting"
              (click)="queue()">
        <mat-icon>add_circle</mat-icon>
        {{ submitting ? 'Queuing...' : 'Queue Job' }}
      </button>
    </mat-dialog-actions>
  `,
  styles: [`
    .loading {
      display: flex;
      justify-content: center;
      padding: 24px;
    }

    .folder-label {
      margin: 0 0 16px;
      font-size: 0.9rem;
      color: #adb5bd;
    }

    .full-width {
      width: 100%;
    }

    .empty-state {
      padding: 24px;
      text-align: center;
      color: #6c757d;
    }

    .history-section {
      margin-top: 16px;
    }

    .history-section h3 {
      font-size: 0.8rem;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: #6c757d;
      margin: 0 0 8px;
    }

    .history-list {
      border: 1px solid #495057;
      border-radius: 4px;
      overflow: hidden;
    }

    .history-item {
      padding: 10px 14px;
      border-bottom: 1px solid #343a40;
      font-size: 0.85rem;
    }

    .history-item:last-child {
      border-bottom: none;
    }

    .history-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 4px;
    }

    .history-date {
      color: #6c757d;
      font-size: 0.8rem;
    }

    .history-path, .history-detail {
      color: #adb5bd;
      font-size: 0.8rem;
      font-family: monospace;
      margin-top: 2px;
      word-break: break-all;
    }

    .history-error {
      color: #f44336;
      font-size: 0.8rem;
      margin-top: 4px;
      padding: 4px 8px;
      background: rgba( 244, 67, 54, 0.1 );
      border-radius: 3px;
    }

    .badge {
      display: inline-block;
      padding: 2px 8px;
      border-radius: 10px;
      font-size: 0.7rem;
      font-weight: 600;
      text-transform: uppercase;
    }

    .badge-pending { background: rgba( 110, 168, 254, 0.15 ); color: #6ea8fe; }
    .badge-processing { background: rgba( 255, 193, 7, 0.15 ); color: #ffc107; }
    .badge-completed { background: rgba( 76, 175, 80, 0.15 ); color: #4caf50; }
    .badge-failed { background: rgba( 244, 67, 54, 0.15 ); color: #f44336; }
    .badge-cancelled { background: rgba( 158, 158, 158, 0.15 ); color: #9e9e9e; }
    .badge-queued { background: rgba( 156, 39, 176, 0.15 ); color: #ce93d8; }

    .path-label {
      color: #6c757d;
    }

    .status-message {
      margin-top: 8px;
      font-size: 0.85rem;
      padding: 8px 12px;
      border-radius: 4px;
    }

    .status-message.error {
      color: #f44336;
      background: rgba( 244, 67, 54, 0.1 );
    }

    .status-message.success {
      color: #4caf50;
      background: rgba( 76, 175, 80, 0.1 );
    }

    h2 {
      display: flex;
      align-items: center;
      gap: 8px;
      cursor: move;
    }
  `]
})
export class QueueJobDialogComponent implements OnInit {

  previousJobs: PreviousJob[] = [];
  videoFiles: VideoFile[] = [];
  selectedVideo = '';
  loading = true;
  submitting = false;
  queued = false;
  error = '';
  successMessage = '';


  constructor(
    public dialogRef: MatDialogRef<QueueJobDialogComponent>,
    @Inject( MAT_DIALOG_DATA ) public data: QueueJobDialogData,
    private http: HttpClient,
  ) {}


  /**
   * Dialog title based on job type.
   */
  get title(): string {
    return this.data.jobType === 'transcription' ? 'Queue Transcription' : 'Queue Claude Processing';
  }


  /**
   * Dialog icon based on job type.
   */
  get icon(): string {
    return this.data.jobType === 'transcription' ? 'subtitles' : 'article';
  }


  /**
   * Whether the queue button should be enabled.
   */
  get canQueue(): boolean {
    if ( this.submitting ) return false;
    if ( this.data.jobType === 'transcription' ) return !!this.selectedVideo;
    return true;
  }


  /**
   * Load previous job history and video files on dialog open.
   */
  ngOnInit(): void {
    let pending = this.data.jobType === 'transcription' ? 2 : 1;
    const done = () => { if ( --pending === 0 ) this.loading = false; };

    this.http.get<PreviousJob[]>( `/api/folders/${ this.data.folderName }/job-history/${ this.data.jobType }` )
      .subscribe({
        next: ( jobs ) => { this.previousJobs = jobs; done(); },
        error: () => done(),
      });

    if ( this.data.jobType === 'transcription' ) {
      this.http.get<VideoFile[]>( `/api/folders/${ this.data.folderName }/video-files` )
        .subscribe({
          next: ( files ) => {
            this.videoFiles = files;
            // Auto-select production MP4 if available, otherwise first file
            const prod = files.find( f => f.filename.includes( '-production' ) );
            this.selectedVideo = prod ? prod.filename : ( files[ 0 ]?.filename || '' );
            done();
          },
          error: () => done(),
        });
    }
  }


  /**
   * Queue the job via the API.
   */
  queue(): void {
    this.submitting = true;
    this.error = '';

    const url = this.data.jobType === 'transcription'
      ? `/api/folders/${ this.data.folderName }/queue-transcription`
      : `/api/folders/${ this.data.folderName }/queue-claude`;

    const body = this.data.jobType === 'transcription'
      ? { video: this.selectedVideo }
      : {};

    this.http.post<{ ok: boolean; jobId: number }>( url, body )
      .subscribe({
        next: ( res ) => {
          this.successMessage = `Job #${ res.jobId } queued`;
          this.submitting = false;
          this.queued = true;
        },
        error: ( err ) => {
          this.error = err.error?.error || 'Failed to queue job';
          this.submitting = false;
        },
      });
  }


  /**
   * Format bytes into a human-readable size.
   *
   * @param bytes - file size in bytes
   * @returns formatted string
   */
  formatSize( bytes: number ): string {
    if ( bytes < 1024 ) return `${ bytes } B`;
    if ( bytes < 1024 * 1024 ) return `${ ( bytes / 1024 ).toFixed( 1 ) } KB`;
    if ( bytes < 1024 * 1024 * 1024 ) return `${ ( bytes / ( 1024 * 1024 ) ).toFixed( 1 ) } MB`;
    return `${ ( bytes / ( 1024 * 1024 * 1024 ) ).toFixed( 2 ) } GB`;
  }


  /**
   * Format a date string for display.
   *
   * @param dateStr - ISO date string
   * @returns formatted date
   */
  formatDate( dateStr: string | null ): string {
    if ( !dateStr ) return '';
    return new Date( dateStr ).toLocaleString( 'en-US', {
      month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit',
    });
  }
}
