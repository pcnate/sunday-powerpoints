import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatIconModule } from '@angular/material/icon';
import { MatSelectModule } from '@angular/material/select';
import { MatButtonModule } from '@angular/material/button';
import { MatDialog, MatDialogModule } from '@angular/material/dialog';
import { AgGridAngular } from 'ag-grid-angular';
import { AllCommunityModule, ColDef, colorSchemeDark, GridApi, GridReadyEvent, ICellRendererParams, ModuleRegistry, RowClickedEvent, themeQuartz, ValueFormatterParams } from 'ag-grid-community';
import { CreateSongDialogComponent } from './create-song-dialog.component';

ModuleRegistry.registerModules([ AllCommunityModule ]);
import { SongDetailDialogComponent, SongDetailDialogData } from './song-detail-dialog.component';


/**
 * Song from the library API.
 */
interface Song {
  name: string;
  number?: string;
  book?: string;
  bookIcon?: string;
  filePath?: string;
  lastModified?: string;
  ccli?: string;
  license?: string;
  lastUsed?: string;
  totalUsed?: number;
}


/**
 * Song Library component — searchable AG Grid table of all available songs.
 */
@Component({
  selector: 'app-library',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatFormFieldModule,
    MatInputModule,
    MatIconModule,
    MatSelectModule,
    MatButtonModule,
    MatDialogModule,
    AgGridAngular,
  ],
  templateUrl: './library.component.html',
  styleUrls: [ './library.component.scss' ]
})
export class LibraryComponent implements OnInit {

  songs: Song[] = [];
  filteredSongs: Song[] = [];
  books: { name: string; count: number; icon?: string }[] = [];
  searchText = '';
  selectedBook = '';
  loading = true;
  displayedRowCount = 0;

  /**
   * AG Grid dark theme matching the app's Bootstrap-dark palette.
   */
  gridTheme = themeQuartz.withPart( colorSchemeDark ).withParams({
    backgroundColor: '#212529',
    headerBackgroundColor: '#1a1d21',
    oddRowBackgroundColor: '#212529',
    rowHoverColor: 'rgba( 110, 168, 254, 0.06 )',
    selectedRowBackgroundColor: 'rgba( 110, 168, 254, 0.12 )',
    borderColor: '#495057',
    headerTextColor: '#ffffff',
    foregroundColor: '#dee2e6',
    rangeSelectionBorderColor: '#6ea8fe',
    fontFamily: 'Roboto, "Helvetica Neue", sans-serif',
    fontSize: 14,
    iconColor: '#adb5bd',
    wrapperBorderRadius: '0px',
    wrapperBorder: false,
  });

  private gridApi!: GridApi;


  /**
   * AG Grid column definitions.
   */
  columnDefs: ColDef[] = [
    {
      field: 'book',
      headerName: 'Book',
      flex: 1,
      minWidth: 120,
      cellRenderer: ( params: ICellRendererParams ) => {
        const name = params.value || '-';
        const icon = params.data?.bookIcon;
        if ( icon ) {
          return `<span class="book-cell"><img src="/api/books/icons/${ icon }" class="book-icon" alt="" />${ name }</span>`;
        }
        return name;
      },
    },
    {
      field: 'number',
      headerName: '#',
      width: 80,
      valueFormatter: ( params: ValueFormatterParams ) => params.value || '-',
    },
    {
      field: 'name',
      headerName: 'Name',
      flex: 2,
      minWidth: 200,
    },
    {
      field: 'ccli',
      headerName: 'CCLI',
      width: 120,
      cellRenderer: ( params: ICellRendererParams ) => {
        if ( !params.value ) return '';
        return `<a class="ccli-link" href="https://songselect.ccli.com/songs/${ params.value }" target="_blank" rel="noopener">${ params.value }</a>`;
      },
      cellClass: ( params ) => params.value ? 'ccli-populated' : 'ccli-missing',
    },
    {
      field: 'license',
      headerName: 'License',
      width: 150,
      valueFormatter: ( params: ValueFormatterParams ) => params.value || '',
    },
    {
      field: 'lastUsed',
      headerName: 'Last Used',
      width: 150,
      valueFormatter: ( params: ValueFormatterParams ) => this.formatDate( params.value ),
      comparator: ( valueA: string, valueB: string ) => {
        const a = padDate( valueA );
        const b = padDate( valueB );
        if ( a === b ) return 0;
        return a < b ? -1 : 1;
      },
    },
    {
      field: 'totalUsed',
      headerName: 'Used',
      width: 90,
      valueFormatter: ( params: ValueFormatterParams ) => params.value ?? 0,
    },
  ];


  /**
   * Default column settings applied to all columns.
   */
  defaultColDef: ColDef = {
    sortable: true,
    resizable: true,
    filter: true,
  };


  /**
   * Overlay template shown when the grid has no rows.
   */
  noRowsTemplate = '<span style="color: #6c757d; font-size: 0.95rem;">No songs found in library</span>';


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
        this.songs = data; // server sorts by book → number → name
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
    const icons = new Map<string, string>();
    for ( const song of this.songs ) {
      const book = song.book || 'Unknown';
      counts.set( book, ( counts.get( book ) || 0 ) + 1 );
      if ( song.bookIcon && !icons.has( book ) ) {
        icons.set( book, song.bookIcon );
      }
    }
    this.books = Array.from( counts.entries() )
      .map( ([ name, count ]) => ({ name, count, icon: icons.get( name ) }) )
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

    // Defer count update to next tick so grid processes the new rowData first
    setTimeout( () => this.updateDisplayedRowCount() );
  }


  /**
   * Format a YYYYMMDD or YYYYMM string into a short readable date.
   *
   * @param dateStr - YYYYMMDD or YYYYMM format, or undefined
   * @returns formatted date like "Feb 2, 2026" or "Feb 2026", or empty string
   */
  formatDate( dateStr?: string ): string {
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
   * Store grid API reference when grid is ready.
   *
   * @param event - grid ready event
   */
  onGridReady( event: GridReadyEvent ): void {
    this.gridApi = event.api;
    this.updateDisplayedRowCount();
  }


  /**
   * Update the displayed row count when AG Grid column filters change.
   */
  onFilterChanged(): void {
    this.updateDisplayedRowCount();
  }


  /**
   * Refresh the displayed row count from the grid API.
   */
  private updateDisplayedRowCount(): void {
    if ( this.gridApi ) {
      this.displayedRowCount = this.gridApi.getDisplayedRowCount();
    } else {
      this.displayedRowCount = this.filteredSongs.length;
    }
  }


  /**
   * Handle row click — open song detail dialog.
   *
   * @param event - AG Grid row clicked event
   */
  onRowClicked( event: RowClickedEvent ): void {
    // Don't open the dialog when the user clicks a link (e.g. CCLI)
    const target = event.event?.target as HTMLElement | undefined;
    if ( target?.tagName === 'A' ) return;

    if ( event.data ) {
      this.openSongDetail( event.data );
    }
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


  /**
   * Open the song detail dialog showing usage history.
   *
   * @param song - the song row that was clicked
   */
  openSongDetail( song: Song ): void {
    const dialogRef = this.dialog.open( SongDetailDialogComponent, {
      width: '550px',
      maxHeight: '80vh',
      data: {
        name: song.name,
        number: song.number,
        book: song.book,
        bookIcon: song.bookIcon,
        filePath: song.filePath,
        ccli: song.ccli,
        lastUsed: song.lastUsed,
        totalUsed: song.totalUsed,
      } as SongDetailDialogData,
    });

    dialogRef.afterClosed().subscribe( ( result: { ccli: string | null; license: string | null } | undefined ) => {
      if ( !result ) return;
      song.ccli = result.ccli || undefined;
      this.gridApi.refreshCells({ columns: [ 'ccli' ] });
    });
  }
}


/**
 * Pad a YYYYMM date to YYYYMMDD for consistent sort ordering.
 * Closing songs use YYYYMM; pad with "32" so they sort after all days in that month.
 *
 * @param dateStr - YYYYMMDD, YYYYMM, or undefined
 * @returns padded date string or empty string
 */
function padDate( dateStr?: string ): string {
  if ( !dateStr ) return '';
  return dateStr.length < 8 ? dateStr + '32' : dateStr;
}
