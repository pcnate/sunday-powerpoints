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


/**
 * Song from the library API.
 */
interface Song {
  name: string;
  number?: string;
  book?: string;
  lastModified?: string;
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

  displayedColumns = [ 'number', 'name', 'book' ];


  constructor( private http: HttpClient ) {}


  /**
   * Load all songs on init.
   */
  ngOnInit(): void {
    this.http.get<Song[]>( '/api/songs' ).subscribe({
      next: ( data ) => {
        this.songs = data.sort( ( a, b ) => a.name.localeCompare( b.name ) );
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
        default: return 0;
      }
    });
  }
}


/**
 * Compare two values for sorting.
 *
 * @param a - first value
 * @param b - second value
 * @param isAsc - ascending order
 * @returns comparison result
 */
function compare( a: string, b: string, isAsc: boolean ): number {
  return ( a < b ? -1 : 1 ) * ( isAsc ? 1 : -1 );
}
