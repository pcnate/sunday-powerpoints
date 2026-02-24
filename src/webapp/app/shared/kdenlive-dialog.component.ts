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
 * Data passed to the Kdenlive dialog.
 */
export interface KdenliveDialogData {
  folderName: string;
}


/**
 * Video file info returned by the API.
 */
interface VideoFile {
  filename: string;
  size: number;
  type: string;
}


/**
 * Kdenlive project creation dialog — lists available video files in a folder's
 * Vids/ directory and lets the user assign OBS (MKV) and Camera (MP4) roles
 * before generating a Kdenlive project file.
 */
@Component({
  selector: 'app-kdenlive-dialog',
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
      <mat-icon>movie_creation</mat-icon>
      Create Kdenlive Project
    </h2>

    <mat-dialog-content>
      <div *ngIf="loading" class="loading">
        <mat-spinner diameter="30"></mat-spinner>
      </div>

      <ng-container *ngIf="!loading && !error">
        <p class="folder-label">{{ data.folderName }}</p>

        <div *ngIf="videoFiles.length === 0" class="empty-state">
          No video files found in Vids/ directory.
        </div>

        <ng-container *ngIf="videoFiles.length > 0">
          <div class="file-list">
            <div class="file-item" *ngFor="let f of videoFiles">
              <mat-icon class="file-icon" [class.mkv]="f.type === 'mkv'" [class.mp4]="f.type === 'mp4'">
                {{ f.type === 'mkv' ? 'desktop_windows' : 'videocam' }}
              </mat-icon>
              <span class="file-name">{{ f.filename }}</span>
              <span class="file-size">{{ formatSize( f.size ) }}</span>
              <span class="file-type">{{ f.type.toUpperCase() }}</span>
            </div>
          </div>

          <mat-form-field appearance="outline" class="role-select">
            <mat-label>OBS Recording (MKV — desktop + 3 audio)</mat-label>
            <mat-select [(ngModel)]="selectedObs">
              <mat-option value="">None (camera only)</mat-option>
              <mat-option *ngFor="let f of videoFiles" [value]="f.filename">
                {{ f.filename }} ({{ formatSize( f.size ) }})
              </mat-option>
            </mat-select>
          </mat-form-field>

          <mat-form-field appearance="outline" class="role-select">
            <mat-label>Camera Recording (MP4 — camera + 1 audio)</mat-label>
            <mat-select [(ngModel)]="selectedCamera">
              <mat-option *ngFor="let f of videoFiles" [value]="f.filename">
                {{ f.filename }} ({{ formatSize( f.size ) }})
              </mat-option>
            </mat-select>
          </mat-form-field>
        </ng-container>
      </ng-container>

      <div *ngIf="error" class="status-message error">{{ error }}</div>
      <div *ngIf="successMessage" class="status-message success">{{ successMessage }}</div>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button (click)="dialogRef.close( created )">
        {{ created ? 'Close' : 'Cancel' }}
      </button>
      <button mat-raised-button color="primary"
              *ngIf="!created"
              [disabled]="!canCreate || creating"
              (click)="create()">
        <mat-icon>movie_creation</mat-icon>
        {{ creating ? 'Creating...' : 'Create Project' }}
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

    .empty-state {
      padding: 24px;
      text-align: center;
      color: #6c757d;
    }

    .file-list {
      margin-bottom: 20px;
      border: 1px solid #495057;
      border-radius: 4px;
      overflow: hidden;
    }

    .file-item {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 10px 14px;
      border-bottom: 1px solid #343a40;
    }

    .file-item:last-child {
      border-bottom: none;
    }

    .file-icon {
      font-size: 20px;
      width: 20px;
      height: 20px;
    }

    .file-icon.mkv { color: #6ea8fe; }
    .file-icon.mp4 { color: #75b798; }

    .file-name {
      flex: 1;
      font-family: monospace;
      font-size: 0.85rem;
    }

    .file-size {
      color: #adb5bd;
      font-size: 0.8rem;
      white-space: nowrap;
    }

    .file-type {
      font-size: 0.75rem;
      font-weight: 600;
      padding: 2px 6px;
      border-radius: 3px;
      background: #343a40;
      color: #e0e0e0;
    }

    .role-select {
      width: 100%;
      margin-bottom: 4px;
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
export class KdenliveDialogComponent implements OnInit {

  videoFiles: VideoFile[] = [];
  selectedObs = '';
  selectedCamera = '';
  loading = true;
  creating = false;
  created = false;
  error = '';
  successMessage = '';


  constructor(
    public dialogRef: MatDialogRef<KdenliveDialogComponent>,
    @Inject( MAT_DIALOG_DATA ) public data: KdenliveDialogData,
    private http: HttpClient
  ) {}


  /**
   * Load video files on dialog open.
   */
  ngOnInit(): void {
    this.http.get<VideoFile[]>( `/api/folders/${ this.data.folderName }/video-files` ).subscribe({
      next: ( files ) => {
        this.videoFiles = files;
        this.autoSelect();
        this.loading = false;
      },
      error: ( err ) => {
        this.error = err.error?.error || 'Failed to load video files';
        this.loading = false;
      },
    });
  }


  /**
   * Auto-select the first MKV as OBS and first MP4 as Camera.
   */
  private autoSelect(): void {
    const mkv = this.videoFiles.find( f => f.type === 'mkv' );
    const mp4 = this.videoFiles.find( f => f.type === 'mp4' );
    if ( mkv ) this.selectedObs = mkv.filename;
    if ( mp4 ) this.selectedCamera = mp4.filename;
  }


  /**
   * Whether at least the camera role is assigned and not creating.
   */
  get canCreate(): boolean {
    return !!this.selectedCamera && !this.creating;
  }


  /**
   * Format bytes into a human-readable size string.
   *
   * @param bytes - file size in bytes
   * @returns formatted string like "1.46 GB"
   */
  formatSize( bytes: number ): string {
    if ( bytes < 1024 ) return `${ bytes } B`;
    if ( bytes < 1024 * 1024 ) return `${ ( bytes / 1024 ).toFixed( 1 ) } KB`;
    if ( bytes < 1024 * 1024 * 1024 ) return `${ ( bytes / ( 1024 * 1024 ) ).toFixed( 1 ) } MB`;
    return `${ ( bytes / ( 1024 * 1024 * 1024 ) ).toFixed( 2 ) } GB`;
  }


  /**
   * Call the API to create the Kdenlive project file.
   */
  create(): void {
    this.creating = true;
    this.error = '';
    this.successMessage = '';

    this.http.post<{ ok: boolean; path: string }>(
      `/api/folders/${ this.data.folderName }/create-kdenlive`,
      { obsFile: this.selectedObs || null, cameraFile: this.selectedCamera }
    ).subscribe({
      next: ( res ) => {
        this.successMessage = `Created ${ res.path }`;
        this.creating = false;
        this.created = true;
      },
      error: ( err ) => {
        this.error = err.error?.error || 'Failed to create Kdenlive project';
        this.creating = false;
      },
    });
  }
}
