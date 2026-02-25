import { Component, Inject, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MatDialogModule, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { Router } from '@angular/router';
import { MatButtonModule } from '@angular/material/button';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatIconModule } from '@angular/material/icon';
import { MatInputModule } from '@angular/material/input';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { MatTooltipModule } from '@angular/material/tooltip';
import { DraggableDialogDirective } from '../shared/draggable-dialog.directive';


export interface SongDetailDialogData {
  name: string;
  number?: string;
  book?: string;
  bookIcon?: string;
  filePath?: string;
  ccli?: string;
  lastUsed?: string;
  totalUsed?: number;
}


interface SongHistory {
  id: number | null;
  name: string;
  number: string | null;
  book: string | null;
  ccli: string | null;
  license: string | null;
  totalUsed: number;
  firstUsed: string | null;
  lastUsed: string | null;
  history: UsageEntry[];
}


interface UsageEntry {
  sundayDate: string;
  slotType: 'song' | 'chorus' | 'closing';
  slotNumber: number;
  slotSuffix: string | null;
}


@Component({
  selector: 'app-song-detail-dialog',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatDialogModule,
    MatButtonModule,
    MatFormFieldModule,
    MatIconModule,
    MatInputModule,
    MatProgressSpinnerModule,
    MatTooltipModule,
    DraggableDialogDirective,
  ],
  template: `
    <h2 mat-dialog-title appDraggableDialog>
      <mat-icon>music_note</mat-icon>
      {{ data.number ? data.number + ' - ' : '' }}{{ data.name }}
    </h2>

    <mat-dialog-content>
      <!-- Song Info Header -->
      <div class="info-grid">
        <div class="info-item" *ngIf="data.book">
          <span class="info-label">Book</span>
          <span class="info-value book-value">
            <img *ngIf="data.bookIcon" [src]="'/api/books/icons/' + data.bookIcon" class="book-icon" alt="" />
            {{ data.book }}
          </span>
        </div>
        <div class="info-item" *ngIf="data.number">
          <span class="info-label">Number</span>
          <span class="info-value">{{ data.number }}</span>
        </div>
      </div>

      <!-- File Path -->
      <div class="file-path" *ngIf="data.filePath" (click)="copyPath()">
        <mat-icon>folder_open</mat-icon>
        <span>{{ data.filePath }}</span>
        <mat-icon class="copy-icon" [matTooltip]="copyTooltip">{{ copyIcon }}</mat-icon>
      </div>

      <!-- CCLI & License -->
      <div class="metadata-fields" *ngIf="!loading">
        <mat-form-field appearance="outline" class="metadata-field">
          <mat-label>CCLI #</mat-label>
          <input matInput [(ngModel)]="ccli" (blur)="saveMetadata()" placeholder="e.g. 1234567" />
          <a *ngIf="ccli" matSuffix class="ccli-link"
             [href]="'https://songselect.ccli.com/songs/' + ccli"
             target="_blank" rel="noopener" (click)="$event.stopPropagation()">
            <mat-icon>open_in_new</mat-icon>
          </a>
        </mat-form-field>
        <mat-form-field appearance="outline" class="metadata-field metadata-field-wide">
          <mat-label>License</mat-label>
          <input matInput [(ngModel)]="license" (blur)="saveMetadata()" placeholder="e.g. CCLI License #12345" />
        </mat-form-field>
        <mat-icon *ngIf="saving" class="save-spinner">sync</mat-icon>
        <mat-icon *ngIf="saved" class="save-check">check_circle</mat-icon>
      </div>

      <!-- Usage Stats -->
      <div class="stats-row">
        <div class="stat">
          <span class="stat-value">{{ songHistory?.totalUsed ?? data.totalUsed ?? 0 }}</span>
          <span class="stat-label">Times Used</span>
        </div>
        <div class="stat" *ngIf="songHistory?.firstUsed">
          <span class="stat-value">{{ formatDate( songHistory!.firstUsed ) }}</span>
          <span class="stat-label">First Used</span>
        </div>
        <div class="stat" *ngIf="songHistory?.lastUsed">
          <span class="stat-value">{{ formatDate( songHistory!.lastUsed ) }}</span>
          <span class="stat-label">Last Used</span>
        </div>
      </div>

      <!-- Loading -->
      <div *ngIf="loading" class="loading">
        <mat-spinner diameter="30"></mat-spinner>
      </div>

      <!-- Usage History List -->
      <div *ngIf="!loading && songHistory?.history?.length" class="history-section">
        <h3 class="history-title">Usage History</h3>
        <div class="history-list">
          <div class="history-item" *ngFor="let entry of songHistory!.history">
            <span class="history-date">{{ formatDate( entry.sundayDate ) }}</span>
            <span class="history-right">
              <span class="history-slot">{{ formatSlot( entry ) }}</span>
              <mat-icon
                class="history-link"
                matTooltip="View in Planning"
                (click)="goToPlanning( entry.sundayDate )">
                open_in_new
              </mat-icon>
            </span>
          </div>
        </div>
      </div>

      <!-- No history -->
      <div *ngIf="!loading && !songHistory?.history?.length" class="no-history">
        <mat-icon>info_outline</mat-icon>
        <span>No usage history recorded</span>
      </div>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button (click)="close()">Close</button>
    </mat-dialog-actions>
  `,
  styles: [`
    h2 {
      display: flex;
      align-items: center;
      gap: 8px;
      cursor: move;
    }

    mat-dialog-content {
      min-width: 400px;
      max-width: 550px;
    }

    .info-grid {
      display: flex;
      gap: 24px;
      margin-bottom: 16px;
      flex-wrap: wrap;
    }

    .info-item {
      display: flex;
      flex-direction: column;
      gap: 2px;
    }

    .info-label {
      font-size: 0.75rem;
      color: #adb5bd;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .info-value {
      font-size: 0.95rem;
      color: #e0e0e0;
    }

    .file-path {
      display: flex;
      align-items: flex-start;
      gap: 6px;
      padding: 8px 10px;
      margin-bottom: 16px;
      background: rgba( 255, 255, 255, 0.04 );
      border-radius: 4px;
      font-size: 0.8rem;
      line-height: 16px;
      color: #adb5bd;
      word-break: break-all;
      user-select: all;
    }

    .file-path mat-icon {
      font-size: 16px;
      width: 16px;
      height: 16px;
      flex-shrink: 0;
      color: #6c757d;
    }

    .file-path .copy-icon {
      margin-left: auto;
      cursor: pointer;
      color: #6ea8fe;
      opacity: 0.6;
      transition: opacity 0.15s;
    }

    .file-path .copy-icon:hover {
      opacity: 1;
    }

    .stats-row {
      display: flex;
      gap: 24px;
      padding: 12px 0;
      border-top: 1px solid #495057;
      border-bottom: 1px solid #495057;
      margin-bottom: 16px;
    }

    .stat {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 2px;
    }

    .stat-value {
      font-size: 1.1rem;
      font-weight: 500;
      color: #6ea8fe;
    }

    .stat-label {
      font-size: 0.75rem;
      color: #adb5bd;
    }

    .loading {
      display: flex;
      justify-content: center;
      padding: 24px;
    }

    .history-title {
      font-size: 0.85rem;
      color: #adb5bd;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin: 0 0 8px 0;
    }

    .history-list {
      max-height: 300px;
      overflow-y: auto;
    }

    .history-item {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 6px 8px;
      border-radius: 4px;
    }

    .history-item:nth-child( even ) {
      background: rgba( 255, 255, 255, 0.03 );
    }

    .history-date {
      color: #e0e0e0;
      font-size: 0.9rem;
    }

    .history-right {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .history-slot {
      color: #adb5bd;
      font-size: 0.85rem;
    }

    .history-link {
      font-size: 16px;
      width: 16px;
      height: 16px;
      cursor: pointer;
      color: #6ea8fe;
      opacity: 0.4;
      transition: opacity 0.15s;
    }

    .history-link:hover {
      opacity: 1;
    }

    .no-history {
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 16px 0;
      color: #6c757d;
      font-size: 0.9rem;
    }

    .no-history mat-icon {
      font-size: 20px;
      width: 20px;
      height: 20px;
    }

    .book-value {
      display: inline-flex;
      align-items: center;
    }

    .book-icon {
      width: 20px;
      height: 20px;
      object-fit: contain;
      margin-right: 6px;
    }

    .metadata-fields {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-bottom: 16px;
    }

    .metadata-field {
      flex: 0 0 140px;
    }

    .metadata-field-wide {
      flex: 1 1 auto;
    }

    ::ng-deep .metadata-field .mat-mdc-form-field-subscript-wrapper {
      display: none;
    }

    .ccli-link {
      color: #6ea8fe;
      display: inline-flex;
      align-items: center;
      text-decoration: none;
    }

    .ccli-link mat-icon {
      font-size: 16px;
      width: 16px;
      height: 16px;
    }

    .ccli-link:hover {
      color: #9ec5fe;
    }

    .save-spinner {
      font-size: 18px;
      width: 18px;
      height: 18px;
      color: #6ea8fe;
      animation: spin 1s linear infinite;
      flex-shrink: 0;
    }

    .save-check {
      font-size: 18px;
      width: 18px;
      height: 18px;
      color: #75b798;
      flex-shrink: 0;
    }

    @keyframes spin {
      100% { transform: rotate( 360deg ); }
    }
  `]
})
export class SongDetailDialogComponent implements OnInit {

  songHistory: SongHistory | null = null;
  loading = true;
  copyIcon = 'content_copy';
  copyTooltip = 'Copy path';

  ccli = '';
  license = '';
  saving = false;
  saved = false;
  private savedTimer: any;
  private metadataChanged = false;


  constructor(
    @Inject( MAT_DIALOG_DATA ) public data: SongDetailDialogData,
    private dialogRef: MatDialogRef<SongDetailDialogComponent>,
    private http: HttpClient,
    private router: Router
  ) {}


  /**
   * Fetch usage history from the server on init.
   */
  ngOnInit(): void {
    const params = new URLSearchParams();
    params.set( 'name', this.data.name );
    if ( this.data.number ) params.set( 'number', this.data.number );
    if ( this.data.book ) params.set( 'book', this.data.book );

    this.http.get<SongHistory>( `/api/songs/history?${ params.toString() }` ).subscribe({
      next: ( result ) => {
        this.songHistory = result;
        this.ccli = result.ccli || '';
        this.license = result.license || '';
        this.loading = false;
      },
      error: () => {
        this.loading = false;
      },
    });
  }


  /**
   * Copy the file path to the clipboard with visual feedback.
   */
  copyPath(): void {
    if ( !this.data.filePath ) return;
    navigator.clipboard.writeText( this.data.filePath ).then( () => {
      this.copyIcon = 'check';
      this.copyTooltip = 'Copied!';
      setTimeout( () => {
        this.copyIcon = 'content_copy';
        this.copyTooltip = 'Copy path';
      }, 2000 );
    });
  }


  /**
   * Save CCLI and license to the server when a field loses focus.
   * If the song has no DB record yet (id is null), passes song identifiers
   * so the server can find-or-create one.
   * Shows a brief checkmark on success.
   */
  saveMetadata(): void {
    // Only save if values actually changed
    const currentCcli = this.songHistory?.ccli || '';
    const currentLicense = this.songHistory?.license || '';
    if ( this.ccli === currentCcli && this.license === currentLicense ) return;

    this.saving = true;
    this.saved = false;
    clearTimeout( this.savedTimer );

    const songId = this.songHistory?.id || 0;
    this.http.put<{ ok: boolean; id?: number }>( `/api/songs/${ songId }`, {
      ccli: this.ccli || null,
      license: this.license || null,
      name: this.data.name,
      number: this.data.number || null,
      book: this.data.book || null,
    }).subscribe({
      next: ( res ) => {
        if ( this.songHistory ) {
          this.songHistory.ccli = this.ccli || null;
          this.songHistory.license = this.license || null;
          if ( res.id && !this.songHistory.id ) {
            this.songHistory.id = res.id;
          }
        }
        this.metadataChanged = true;
        this.saving = false;
        this.saved = true;
        this.savedTimer = setTimeout( () => { this.saved = false; }, 2000 );
      },
      error: () => {
        this.saving = false;
      },
    });
  }


  /**
   * Close the dialog, returning updated metadata if it was changed.
   */
  close(): void {
    if ( this.metadataChanged ) {
      this.dialogRef.close({ ccli: this.ccli || null, license: this.license || null });
    } else {
      this.dialogRef.close();
    }
  }


  /**
   * Format a YYYYMMDD or YYYYMM string into a readable date.
   *
   * @param dateStr - YYYYMMDD or YYYYMM format, or null
   * @returns formatted date like "Feb 2, 2026" or "Feb 2026", or empty string
   */
  formatDate( dateStr?: string | null ): string {
    if ( !dateStr || dateStr.length < 6 ) return '';
    const y = parseInt( dateStr.slice( 0, 4 ) );
    const m = parseInt( dateStr.slice( 4, 6 ) ) - 1;
    if ( dateStr.length < 8 ) {
      return new Date( y, m, 1 ).toLocaleDateString( 'en-US', { month: 'short', year: 'numeric' });
    }
    const d = parseInt( dateStr.slice( 6, 8 ) );
    return new Date( y, m, d ).toLocaleDateString( 'en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }


  /**
   * Format a usage entry's slot type and number into a readable label.
   *
   * @param entry - the usage history entry
   * @returns label like "Song 1", "Chorus 2", or "Closing Song"
   */
  formatSlot( entry: UsageEntry ): string {
    if ( entry.slotType === 'closing' ) return 'Closing Song';
    const type = entry.slotType === 'song' ? 'Song' : 'Chorus';
    const suffix = entry.slotSuffix ? entry.slotSuffix : '';
    return `${ type } ${ entry.slotNumber }${ suffix }`;
  }


  /**
   * Close the dialog and navigate to the Planning tab for the given date's month.
   *
   * @param dateStr - YYYYMMDD format sunday date
   */
  goToPlanning( dateStr: string ): void {
    const year = parseInt( dateStr.slice( 0, 4 ) );
    const month = parseInt( dateStr.slice( 4, 6 ) );
    this.dialogRef.close();
    this.router.navigate( [ '/planning', year, month ] );
  }
}
