import { Component, Inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MatDialogModule, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { DraggableDialogDirective } from '../shared/draggable-dialog.directive';


/**
 * Data passed into the YouTube URL dialog.
 */
export interface YoutubeDialogData {
  folder: string;
  date: string;
  currentUrl: string | null;
}


/**
 * YouTube URL editor dialog — set or update the YouTube video link for a Sunday.
 */
@Component( {
  selector: 'app-youtube-dialog',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatDialogModule,
    MatButtonModule,
    MatIconModule,
    MatFormFieldModule,
    MatInputModule,
    DraggableDialogDirective,
  ],
  template: `
    <h2 mat-dialog-title appDraggableDialog>
      <mat-icon>smart_display</mat-icon>
      YouTube URL — {{ data.date }}
    </h2>

    <mat-dialog-content>
      <mat-form-field appearance="outline" class="full-width">
        <mat-label>YouTube URL</mat-label>
        <input matInput [(ngModel)]="url" placeholder="https://www.youtube.com/watch?v=...">
      </mat-form-field>

      <a *ngIf="url?.trim()" [href]="url" target="_blank" rel="noopener"
         mat-button class="preview-link">
        <mat-icon>open_in_new</mat-icon>
        Open in YouTube
      </a>

      <div *ngIf="statusMessage" class="status-message" [class.error]="hasError">
        {{ statusMessage }}
      </div>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button (click)="dialogRef.close()">Cancel</button>
      <button mat-button color="warn" *ngIf="data.currentUrl" (click)="remove()" [disabled]="saving">
        <mat-icon>delete</mat-icon>
        Remove
      </button>
      <button mat-raised-button color="primary" (click)="save()" [disabled]="saving || !url?.trim()">
        <mat-icon>save</mat-icon>
        Save
      </button>
    </mat-dialog-actions>
  `,
  styles: [`
    .full-width {
      width: 100%;
    }

    .preview-link {
      color: #6ea8fe;
      margin-bottom: 12px;
      display: inline-flex;
      align-items: center;
      gap: 4px;
    }

    .status-message {
      margin-top: 8px;
      font-size: 0.85rem;
      color: #4caf50;
    }

    .status-message.error {
      color: #f44336;
    }

    h2 {
      display: flex;
      align-items: center;
      gap: 8px;
      cursor: move;
    }

    mat-dialog-content {
      min-width: 400px;
      overflow: visible;
    }
  `]
} )
export class YoutubeDialogComponent {

  url: string;
  saving = false;
  statusMessage = '';
  hasError = false;


  constructor(
    public dialogRef: MatDialogRef<YoutubeDialogComponent>,
    @Inject( MAT_DIALOG_DATA ) public data: YoutubeDialogData,
    private http: HttpClient
  ) {
    this.url = data.currentUrl || '';
  }


  /**
   * Save the YouTube URL to the server, then close.
   */
  save(): void {
    this.saving = true;
    this.statusMessage = '';
    this.hasError = false;

    this.http.put( '/api/youtube-url', {
      folder: this.data.folder,
      url: this.url.trim(),
    } ).subscribe( {
      next: () => {
        this.saving = false;
        this.dialogRef.close( { url: this.url.trim() } );
      },
      error: ( err ) => {
        this.statusMessage = `Save failed: ${ err.message || 'Unknown error' }`;
        this.hasError = true;
        this.saving = false;
      },
    } );
  }


  /**
   * Remove the YouTube URL and close.
   */
  remove(): void {
    this.saving = true;
    this.statusMessage = '';
    this.hasError = false;

    this.http.put( '/api/youtube-url', {
      folder: this.data.folder,
      url: '',
    } ).subscribe( {
      next: () => {
        this.saving = false;
        this.dialogRef.close( { url: null } );
      },
      error: ( err ) => {
        this.statusMessage = `Remove failed: ${ err.message || 'Unknown error' }`;
        this.hasError = true;
        this.saving = false;
      },
    } );
  }
}
