import { Component, Inject, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MAT_DIALOG_DATA, MatDialogModule, MatDialogRef } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatListModule } from '@angular/material/list';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';


/**
 * Data passed to the file browser dialog.
 */
export interface FileBrowserDialogData {
  title: string;
  startPath: string;
  extensions?: string[];
  directoryMode?: boolean;
}


/**
 * A single entry returned from the browse API.
 */
interface BrowseEntry {
  name: string;
  is_dir: boolean;
  size: number;
}


/**
 * Response from POST /api/browse.
 */
interface BrowseResponse {
  path: string;
  entries: BrowseEntry[];
}


/**
 * File browser dialog for navigating the worker's filesystem.
 * Replaces the native rfd file dialog with a web-based alternative.
 */
@Component({
  selector: 'app-file-browser-dialog',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatDialogModule,
    MatButtonModule,
    MatIconModule,
    MatFormFieldModule,
    MatInputModule,
    MatListModule,
    MatProgressSpinnerModule,
  ],
  template: `
    <h2 mat-dialog-title>
      <mat-icon>folder_open</mat-icon>
      {{ data.title }}
    </h2>

    <mat-dialog-content>
      <!-- Path bar -->
      <div class="path-bar">
        <mat-form-field appearance="outline" class="path-input">
          <input matInput [(ngModel)]="pathInput"
                 (keyup.enter)="navigateTo( pathInput )"
                 placeholder="Path">
        </mat-form-field>
        <button mat-icon-button (click)="navigateTo( pathInput )" matTooltip="Go">
          <mat-icon>arrow_forward</mat-icon>
        </button>
      </div>

      <!-- Error message -->
      <div class="error-msg" *ngIf="error">{{ error }}</div>

      <!-- Loading spinner -->
      <div class="loading" *ngIf="loading">
        <mat-spinner diameter="32"></mat-spinner>
      </div>

      <!-- Entry list -->
      <div class="entry-list" *ngIf="!loading">
        <!-- Go up -->
        <div class="entry entry-dir"
             *ngIf="currentPath"
             (click)="goUp()"
             (dblclick)="goUp()">
          <mat-icon class="entry-icon icon-folder">arrow_upward</mat-icon>
          <span class="entry-name">..</span>
        </div>

        <div *ngFor="let entry of entries"
             class="entry"
             [class.entry-dir]="entry.is_dir"
             [class.entry-file]="!entry.is_dir"
             [class.entry-selected]="selectedEntry === entry"
             (click)="selectEntry( entry )"
             (dblclick)="openEntry( entry )">
          <mat-icon class="entry-icon" [class.icon-folder]="entry.is_dir" [class.icon-file]="!entry.is_dir">
            {{ entry.is_dir ? 'folder' : 'insert_drive_file' }}
          </mat-icon>
          <span class="entry-name">{{ entry.name }}</span>
          <span class="entry-size" *ngIf="!entry.is_dir">{{ formatSize( entry.size ) }}</span>
        </div>

        <div class="empty-state" *ngIf="entries.length === 0 && !error">
          No matching files
        </div>
      </div>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button (click)="cancel()">Cancel</button>
      <button mat-raised-button color="primary"
              [disabled]="!data.directoryMode && (!selectedEntry || selectedEntry.is_dir)"
              (click)="confirm()">
        {{ data.directoryMode ? 'Select Directory' : 'Select' }}
      </button>
    </mat-dialog-actions>
  `,
  styles: [`
    :host {
      display: block;
    }

    h2[mat-dialog-title] {
      display: flex;
      align-items: center;
      gap: 8px;
      color: #cdd6f4;
      margin: 0;
      padding: 16px 24px;
    }

    mat-dialog-content {
      min-height: 350px;
      max-height: 450px;
      padding: 0 24px;
    }

    .path-bar {
      display: flex;
      align-items: center;
      gap: 4px;
      margin-bottom: 8px;
    }

    .path-input {
      flex: 1;
    }

    .error-msg {
      padding: 8px 12px;
      margin-bottom: 8px;
      border-radius: 6px;
      background: rgba(243, 139, 168, 0.15);
      color: #f38ba8;
      font-size: 0.85rem;
    }

    .loading {
      display: flex;
      justify-content: center;
      padding: 48px 0;
    }

    .entry-list {
      overflow-y: auto;
      max-height: 340px;
    }

    .entry {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 6px 12px;
      border-radius: 6px;
      cursor: pointer;
      user-select: none;
      font-size: 0.9rem;
    }

    .entry:hover {
      background: rgba(137, 180, 250, 0.08);
    }

    .entry-selected {
      background: rgba(137, 180, 250, 0.15) !important;
    }

    .entry-icon {
      font-size: 20px;
      width: 20px;
      height: 20px;
      flex-shrink: 0;
    }

    .icon-folder {
      color: #f9e2af;
    }

    .icon-file {
      color: #a6adc8;
    }

    .entry-name {
      flex: 1;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .entry-size {
      font-size: 0.8rem;
      color: #a6adc8;
      flex-shrink: 0;
    }

    .empty-state {
      text-align: center;
      padding: 32px;
      color: #a6adc8;
      font-size: 0.9rem;
    }

    mat-dialog-actions {
      padding: 12px 24px;
    }
  `]
})
export class FileBrowserDialogComponent implements OnInit {

  currentPath = '';
  pathInput = '';
  entries: BrowseEntry[] = [];
  selectedEntry: BrowseEntry | null = null;
  loading = false;
  error = '';


  constructor(
    private http: HttpClient,
    private dialogRef: MatDialogRef<FileBrowserDialogComponent>,
    @Inject( MAT_DIALOG_DATA ) public data: FileBrowserDialogData,
  ) {}


  /**
   * Load the initial directory on open.
   */
  ngOnInit(): void {
    if ( this.data.startPath ) {
      // Navigate to parent directory of the current value
      const parent = this.getParentPath( this.data.startPath );
      this.navigateTo( parent || '' );
    } else {
      this.navigateTo( '' );
    }
  }


  /**
   * Fetch directory listing from the worker API.
   *
   * @param path - directory to list, empty for drive listing
   */
  navigateTo( path: string ): void {
    this.loading = true;
    this.error = '';
    this.selectedEntry = null;

    this.http.post<BrowseResponse>( '/api/browse', {
      path: path || undefined,
      extensions: this.data.extensions,
    } ).subscribe({
      next: ( res ) => {
        this.loading = false;
        this.currentPath = res.path;
        this.pathInput = res.path;
        this.entries = res.entries;
      },
      error: ( err ) => {
        this.loading = false;
        this.error = err.error?.error || 'Failed to browse directory';
      },
    });
  }


  /**
   * Navigate to the parent directory, or back to drive listing from a drive root.
   */
  goUp(): void {
    if ( !this.currentPath ) return;
    const parent = this.getParentPath( this.currentPath );
    this.navigateTo( parent || '' );
  }


  /**
   * Select an entry (single click).
   *
   * @param entry - the clicked entry
   */
  selectEntry( entry: BrowseEntry ): void {
    this.selectedEntry = entry;
  }


  /**
   * Open an entry (double click): navigate into directories, confirm files.
   *
   * @param entry - the double-clicked entry
   */
  openEntry( entry: BrowseEntry ): void {
    if ( entry.is_dir ) {
      // Drive listing (currentPath empty) → entry.name is already "C:\"
      if ( !this.currentPath ) {
        this.navigateTo( entry.name );
      } else {
        const sep = this.currentPath.endsWith( '\\' ) ? '' : '\\';
        this.navigateTo( this.currentPath + sep + entry.name );
      }
    } else {
      this.selectedEntry = entry;
      this.confirm();
    }
  }


  /**
   * Close dialog with the selected file/directory path.
   */
  confirm(): void {
    if ( this.data.directoryMode ) {
      // In directory mode: select a highlighted directory, or the current directory
      if ( this.selectedEntry && this.selectedEntry.is_dir ) {
        const sep = this.currentPath.endsWith( '\\' ) ? '' : '\\';
        this.dialogRef.close( this.currentPath + sep + this.selectedEntry.name );
      } else {
        this.dialogRef.close( this.currentPath );
      }
      return;
    }
    if ( !this.selectedEntry || this.selectedEntry.is_dir ) return;
    const sep = this.currentPath.endsWith( '\\' ) ? '' : '\\';
    this.dialogRef.close( this.currentPath + sep + this.selectedEntry.name );
  }


  /**
   * Close dialog with no selection.
   */
  cancel(): void {
    this.dialogRef.close( null );
  }


  /**
   * Format byte count as a human-readable size string.
   *
   * @param bytes - file size in bytes
   * @returns formatted string like "1.5 MB"
   */
  formatSize( bytes: number ): string {
    if ( bytes < 1024 ) return bytes + ' B';
    if ( bytes < 1024 * 1024 ) return ( bytes / 1024 ).toFixed( 1 ) + ' KB';
    if ( bytes < 1024 * 1024 * 1024 ) return ( bytes / ( 1024 * 1024 ) ).toFixed( 1 ) + ' MB';
    return ( bytes / ( 1024 * 1024 * 1024 ) ).toFixed( 1 ) + ' GB';
  }


  /**
   * Get the parent path of a given path.
   * Returns empty string for drive roots (e.g. "C:\") or non-path values.
   *
   * @param path - filesystem path
   * @returns parent path or empty string
   */
  private getParentPath( path: string ): string {
    // Not a real filesystem path (no backslash) → drive listing
    if ( !path.includes( '\\' ) ) return '';
    // Drive root like "C:\" → go to drive listing
    if ( /^[A-Za-z]:\\?$/.test( path ) ) return '';
    // Remove trailing backslash then find last separator
    const normalized = path.replace( /\\$/, '' );
    const lastSep = normalized.lastIndexOf( '\\' );
    if ( lastSep <= 2 ) {
      // Parent is the drive root
      return normalized.substring( 0, 3 );
    }
    return normalized.substring( 0, lastSep );
  }
}
