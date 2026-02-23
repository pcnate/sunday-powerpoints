import { Component, Inject, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MAT_DIALOG_DATA, MatDialog, MatDialogModule, MatDialogRef } from '@angular/material/dialog';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatSelectModule } from '@angular/material/select';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatTableModule } from '@angular/material/table';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { DraggableDialogDirective } from '../shared/draggable-dialog.directive';
import { CreateSongDialogComponent } from '../library/create-song-dialog.component';


/**
 * A song from the library API.
 */
interface LibrarySong {
  name: string;
  number?: string;
  book?: string;
  lastModified?: string;
  lastUsed?: string;
  totalUsed?: number;
}


/**
 * Data passed into the song select dialog.
 */
interface DialogData {
  currentSong: { name: string; number?: string } | null;
  slot: number | string;
}


/**
 * Dialog for selecting a song from the library.
 * Provides search by name/number, book/source dropdown filter,
 * and a table with Select buttons per row.
 */
@Component({
  selector: 'app-song-select-dialog',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatDialogModule,
    MatFormFieldModule,
    MatInputModule,
    MatSelectModule,
    MatButtonModule,
    MatIconModule,
    MatTableModule,
    MatProgressSpinnerModule,
    DraggableDialogDirective,
  ],
  template: `
    <h2 mat-dialog-title appDraggableDialog>
      {{ data.slot === 'chorus' ? 'Select Chorus' : 'Select Song ' + data.slot }}
    </h2>

    <mat-dialog-content>
      <div class="filter-row">
        <mat-form-field appearance="outline" class="search-field">
          <mat-label>Search by number or name...</mat-label>
          <input matInput [(ngModel)]="searchText" (ngModelChange)="applyFilters()"
                 placeholder="Search by number or name...">
          <mat-icon matSuffix>search</mat-icon>
        </mat-form-field>

        <mat-form-field appearance="outline" class="book-field">
          <mat-label>Book/Source</mat-label>
          <mat-select [(ngModel)]="selectedBook" (selectionChange)="applyFilters()">
            <mat-option value="">All Books/Sources ({{ songs.length }})</mat-option>
            <mat-option *ngFor="let b of books" [value]="b.name">
              {{ b.name }} ({{ b.count }})
            </mat-option>
          </mat-select>
        </mat-form-field>
      </div>

      <div class="results-count">Results: {{ filteredSongs.length }}</div>

      <div *ngIf="loading" class="loading-state">
        <mat-spinner diameter="36"></mat-spinner>
      </div>

      <div class="table-container" *ngIf="!loading">
        <table mat-table [dataSource]="filteredSongs" class="select-table">

          <ng-container matColumnDef="number">
            <th mat-header-cell *matHeaderCellDef>Number</th>
            <td mat-cell *matCellDef="let song">{{ song.number || '' }}</td>
          </ng-container>

          <ng-container matColumnDef="name">
            <th mat-header-cell *matHeaderCellDef>Name</th>
            <td mat-cell *matCellDef="let song">{{ song.name }}</td>
          </ng-container>

          <ng-container matColumnDef="book">
            <th mat-header-cell *matHeaderCellDef>Book</th>
            <td mat-cell *matCellDef="let song">{{ song.book || '' }}</td>
          </ng-container>

          <ng-container matColumnDef="lastUsed">
            <th mat-header-cell *matHeaderCellDef>Last Used</th>
            <td mat-cell *matCellDef="let song">{{ formatDate( song.lastUsed ) }}</td>
          </ng-container>

          <ng-container matColumnDef="totalUsed">
            <th mat-header-cell *matHeaderCellDef>Used</th>
            <td mat-cell *matCellDef="let song">{{ song.totalUsed ?? 0 }}</td>
          </ng-container>

          <ng-container matColumnDef="action">
            <th mat-header-cell *matHeaderCellDef></th>
            <td mat-cell *matCellDef="let song">
              <button mat-raised-button color="primary" (click)="select( song )">Select</button>
            </td>
          </ng-container>

          <tr mat-header-row *matHeaderRowDef="displayedColumns; sticky: true"></tr>
          <tr mat-row *matRowDef="let row; columns: displayedColumns;"
              (click)="select( row )" class="clickable-row"></tr>
        </table>

        <div *ngIf="filteredSongs.length === 0" class="no-results">
          No songs match your search
        </div>
      </div>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button mat-dialog-close>Cancel</button>
      <button mat-button color="warn" *ngIf="data.currentSong" (click)="clear()">Clear Slot</button>
      <span class="action-spacer"></span>
      <button mat-stroked-button (click)="createNew()">
        <mat-icon>add</mat-icon> Create New
      </button>
    </mat-dialog-actions>
  `,
  styles: [`
    .filter-row {
      display: flex;
      gap: 12px;
      margin-bottom: 4px;
    }

    .search-field { flex: 1; }

    .book-field { min-width: 220px; }

    .results-count {
      font-size: 0.85rem;
      color: #6c757d;
      margin-bottom: 8px;
    }

    mat-dialog-content {
      min-width: 700px;
      max-width: 900px;
      overflow: visible;
    }

    .table-container {
      max-height: 400px;
      overflow-y: auto;
      border: 1px solid #495057;
      border-radius: 4px;
    }

    .select-table {
      width: 100%;
    }

    .clickable-row {
      cursor: pointer;
    }

    .clickable-row:hover {
      background: rgba( 110, 168, 254, 0.06 );
    }

    .loading-state {
      text-align: center;
      padding: 24px;
    }

    .no-results {
      text-align: center;
      padding: 24px;
      color: #6c757d;
    }

    .action-spacer {
      flex: 1;
    }
  `]
})
export class SongSelectDialogComponent implements OnInit {

  songs: LibrarySong[] = [];
  filteredSongs: LibrarySong[] = [];
  books: { name: string; count: number }[] = [];
  searchText = '';
  selectedBook = '';
  loading = true;

  displayedColumns = [ 'number', 'name', 'book', 'lastUsed', 'totalUsed', 'action' ];


  constructor(
    private http: HttpClient,
    private dialog: MatDialog,
    private dialogRef: MatDialogRef<SongSelectDialogComponent>,
    @Inject( MAT_DIALOG_DATA ) public data: DialogData
  ) {}


  /**
   * Load the song library on init.
   */
  ngOnInit(): void {
    this.http.get<LibrarySong[]>( '/api/songs' ).subscribe({
      next: ( songs ) => {
        this.songs = songs.sort( ( a, b ) => {
          if ( a.number && b.number ) return a.number.localeCompare( b.number, undefined, { numeric: true } );
          if ( a.number ) return -1;
          if ( b.number ) return 1;
          return a.name.localeCompare( b.name );
        });
        this.filteredSongs = this.songs;
        this.buildBookList();
        this.loading = false;
      },
      error: () => { this.loading = false; },
    });
  }


  /**
   * Build the list of unique books/sources with counts.
   */
  private buildBookList(): void {
    const counts = new Map<string, number>();
    for ( const song of this.songs ) {
      const book = song.book || 'Unknown';
      counts.set( book, ( counts.get( book ) || 0 ) + 1 );
    }
    this.books = Array.from( counts.entries() )
      .map( ([ name, count ]) => ({ name, count }) )
      .sort( ( a, b ) => a.name.localeCompare( b.name ) );
  }


  /**
   * Apply both search text and book filter.
   */
  applyFilters(): void {
    const q = this.searchText.toLowerCase().trim();

    this.filteredSongs = this.songs.filter( s => {
      // Book filter
      if ( this.selectedBook && ( s.book || 'Unknown' ) !== this.selectedBook ) {
        return false;
      }
      // Text search
      if ( q ) {
        return s.name.toLowerCase().includes( q ) ||
          ( s.number && s.number.includes( q ) ) ||
          ( s.book && s.book.toLowerCase().includes( q ) );
      }
      return true;
    });
  }


  /**
   * Select a song and close the dialog.
   *
   * @param song - the chosen song
   */
  select( song: LibrarySong ): void {
    this.dialogRef.close({ name: song.name, number: song.number, book: song.book });
  }


  /**
   * Clear the slot by returning an empty object.
   */
  clear(): void {
    this.dialogRef.close({ name: '', number: '' });
  }


  /**
   * Format a YYYYMMDD date string into a readable date.
   *
   * @param dateStr - date in YYYYMMDD format
   * @returns formatted date string, or empty string if not provided
   */
  formatDate( dateStr?: string ): string {
    if ( !dateStr || dateStr.length !== 8 ) return '';
    const y = parseInt( dateStr.slice( 0, 4 ) );
    const m = parseInt( dateStr.slice( 4, 6 ) ) - 1;
    const d = parseInt( dateStr.slice( 6, 8 ) );
    return new Date( y, m, d ).toLocaleDateString( 'en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }


  /**
   * Open the create song dialog. On success, auto-select the newly created song.
   */
  createNew(): void {
    const createRef = this.dialog.open( CreateSongDialogComponent, {
      width: '550px',
    });

    createRef.afterClosed().subscribe( ( result: { name: string; number?: string; book?: string } | undefined ) => {
      if ( result ) {
        this.dialogRef.close({ name: result.name, number: result.number, book: result.book });
      }
    });
  }
}
