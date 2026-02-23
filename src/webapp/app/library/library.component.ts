import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MatTableModule } from '@angular/material/table';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatIconModule } from '@angular/material/icon';
import { MatSortModule, Sort } from '@angular/material/sort';
import { MatSelectModule } from '@angular/material/select';
import { MatButtonModule } from '@angular/material/button';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { MatDialog, MatDialogModule } from '@angular/material/dialog';
import { CreateSongDialogComponent } from './create-song-dialog.component';


/**
 * Song from the library API.
 */
interface Song {
  name: string;
  number?: string;
  book?: string;
  lastModified?: string;
  ccli?: string;
  lastUsed?: string;
  totalUsed?: number;
}


/**
 * Song Library component — searchable table of all available songs.
 */
@Component({
  selector: 'app-library',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatTableModule,
    MatFormFieldModule,
    MatInputModule,
    MatIconModule,
    MatSortModule,
    MatSelectModule,
    MatButtonModule,
    MatProgressSpinnerModule,
    MatDialogModule,
  ],
  templateUrl: './library.component.html',
  styleUrls: [ './library.component.scss' ]
})
export class LibraryComponent implements OnInit {

  songs: Song[] = [];
  filteredSongs: Song[] = [];
  books: { name: string; count: number }[] = [];
  searchText = '';
  selectedBook = '';
  loading = true;

  displayedColumns = [ 'number', 'name', 'book', 'ccli', 'lastUsed', 'totalUsed' ];


  constructor(
    private http: HttpClient,
    private dialog: MatDialog
  ) {}


  /**
   * Load all songs on init.
   */
  ngOnInit(): void {
    this.loadSongs();
  }


  /**
   * Fetch the song list from the server.
   */
  private loadSongs(): void {
    this.loading = true;
    this.http.get<Song[]>( '/api/songs' ).subscribe({
      next: ( data ) => {
        this.songs = data.sort( ( a, b ) => a.name.localeCompare( b.name ) );
        this.filteredSongs = this.songs;
        this.buildBookList();
        this.filterSongs();
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
   * Filter songs by search text and book filter.
   */
  filterSongs(): void {
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
   * Sort the table by column.
   *
   * @param sort - the sort event from MatSort
   */
  sortData( sort: Sort ): void {
    if ( !sort.active || sort.direction === '' ) {
      this.filteredSongs = [ ...this.filteredSongs ];
      return;
    }

    this.filteredSongs = [ ...this.filteredSongs ].sort( ( a, b ) => {
      const isAsc = sort.direction === 'asc';
      switch ( sort.active ) {
        case 'number': return compare( a.number || '', b.number || '', isAsc );
        case 'name': return compare( a.name, b.name, isAsc );
        case 'book': return compare( a.book || '', b.book || '', isAsc );
        case 'ccli': return compare( a.ccli || '', b.ccli || '', isAsc );
        case 'lastUsed': return compare( a.lastUsed || '', b.lastUsed || '', isAsc );
        case 'totalUsed': return compareNum( a.totalUsed || 0, b.totalUsed || 0, isAsc );
        default: return 0;
      }
    });
  }


  /**
   * Format a YYYYMMDD string into a short readable date.
   *
   * @param dateStr - YYYYMMDD format or undefined
   * @returns formatted date like "Feb 2, 2026" or empty string
   */
  formatDate( dateStr?: string ): string {
    if ( !dateStr || dateStr.length < 8 ) return '';
    const y = parseInt( dateStr.slice( 0, 4 ) );
    const m = parseInt( dateStr.slice( 4, 6 ) ) - 1;
    const d = parseInt( dateStr.slice( 6, 8 ) );
    return new Date( y, m, d ).toLocaleDateString( 'en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  }


  /**
   * Open the create song dialog and reload songs on success.
   */
  openCreateDialog(): void {
    const dialogRef = this.dialog.open( CreateSongDialogComponent, {
      width: '550px',
    });

    dialogRef.afterClosed().subscribe( ( result ) => {
      if ( result ) {
        this.loadSongs();
      }
    });
  }
}


/**
 * Compare two string values for sorting.
 *
 * @param a - first value
 * @param b - second value
 * @param isAsc - ascending order
 * @returns comparison result
 */
function compare( a: string, b: string, isAsc: boolean ): number {
  return ( a < b ? -1 : 1 ) * ( isAsc ? 1 : -1 );
}


/**
 * Compare two numeric values for sorting.
 *
 * @param a - first value
 * @param b - second value
 * @param isAsc - ascending order
 * @returns comparison result
 */
function compareNum( a: number, b: number, isAsc: boolean ): number {
  return ( a - b ) * ( isAsc ? 1 : -1 );
}
