import { Component, Inject, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { MatDialogModule, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { DraggableDialogDirective } from '../shared/draggable-dialog.directive';


export interface SongDetailDialogData {
  name: string;
  number?: string;
  book?: string;
  ccli?: string;
  lastUsed?: string;
  totalUsed?: number;
}


interface SongHistory {
  name: string;
  number: string | null;
  book: string | null;
  ccli: string | null;
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
    MatDialogModule,
    MatButtonModule,
    MatIconModule,
    MatProgressSpinnerModule,
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
          <span class="info-value">{{ data.book }}</span>
        </div>
        <div class="info-item" *ngIf="data.number">
          <span class="info-label">Number</span>
          <span class="info-value">{{ data.number }}</span>
        </div>
        <div class="info-item" *ngIf="songHistory?.ccli">
          <span class="info-label">CCLI</span>
          <span class="info-value">{{ songHistory!.ccli }}</span>
        </div>
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
            <span class="history-slot">{{ formatSlot( entry ) }}</span>
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
      <button mat-button mat-dialog-close>Close</button>
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

    .history-slot {
      color: #adb5bd;
      font-size: 0.85rem;
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
  `]
})
export class SongDetailDialogComponent implements OnInit {

  songHistory: SongHistory | null = null;
  loading = true;


  constructor(
    @Inject( MAT_DIALOG_DATA ) public data: SongDetailDialogData,
    private http: HttpClient
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
        this.loading = false;
      },
      error: () => {
        this.loading = false;
      },
    });
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
}
