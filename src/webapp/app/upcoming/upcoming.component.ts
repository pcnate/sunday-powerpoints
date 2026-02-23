import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { MatCardModule } from '@angular/material/card';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatChipsModule } from '@angular/material/chips';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatSelectModule } from '@angular/material/select';
import { MatTooltipModule } from '@angular/material/tooltip';
import { MatDialogModule, MatDialog } from '@angular/material/dialog';
import { MatDividerModule } from '@angular/material/divider';
import { MatMenuModule } from '@angular/material/menu';
import { FormsModule } from '@angular/forms';
import { Subject, takeUntil, debounceTime } from 'rxjs';
import { SocketService } from '../../socket.service';
import { SongSelectDialogComponent } from './song-select-dialog.component';
import { NotesDialogComponent } from '../shared/notes-dialog.component';
import { VerseDialogComponent, VerseDialogData, VerseDialogResult } from './verse-dialog.component';
import { ConfirmDialogComponent } from '../admin/confirm-dialog.component';
import { ChorusEditDialogComponent, ChorusEditDialogResult } from './chorus-edit-dialog.component';
import { YoutubeDialogComponent, YoutubeDialogData } from './youtube-dialog.component';


/**
 * Song object from the API.
 */
interface Song {
  name: string;
  number?: string;
  shortcut?: string;
}


/**
 * Week selection data from the API.
 */
interface WeekSelection {
  songs: ( Song | null )[];
  choruses: Song[];
}


/**
 * Folder metadata from the API.
 */
interface FolderInfo {
  name: string;
  hasPre: boolean;
  isApproved: boolean;
  presentationName: string | null;
  presentationSize: string | null;
  presentationModified: string | null;
  hasMp4: boolean;
  videoFileName: string | null;
  youtubeUrl: string | null;
  thumbnail: string | null;
  hasSong1: boolean;
  hasSong2: boolean;
  hasSong3: boolean;
  hasNotes: string | null;
}


/**
 * Upcoming Songs component — the primary weekly workflow.
 * Shows Sundays for the selected month with song/chorus assignments.
 */
@Component({
  selector: 'app-upcoming',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatCardModule,
    MatButtonModule,
    MatIconModule,
    MatChipsModule,
    MatFormFieldModule,
    MatInputModule,
    MatSelectModule,
    MatTooltipModule,
    MatDialogModule,
    MatDividerModule,
    MatMenuModule,
  ],
  templateUrl: './upcoming.component.html',
  styleUrls: [ './upcoming.component.scss' ]
})
export class UpcomingComponent implements OnInit, OnDestroy {

  selectedYear: number;
  selectedMonth: number;
  months = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  years: number[] = [];

  weeks: { date: string; label: string; selection: WeekSelection | null; folder: FolderInfo | null }[] = [];
  folders: FolderInfo[] = [];
  /** True while the initial data fetch for the current month hasn't completed yet. */
  loading = true;

  verseReference = '';
  verseText = '';
  closingSong: Song | null = null;
  templateInfo: { name: string; lastModified: string; size: string } | null = null;

  /** Set of week dates (YYYYMMDD) that the user has unlocked for editing despite being past. */
  editingPastWeeks = new Set<string>();

  private destroy$ = new Subject<void>();
  private reload$ = new Subject<void>();


  constructor(
    private http: HttpClient,
    private dialog: MatDialog,
    private socketService: SocketService
  ) {
    const [ defaultMonth, defaultYear ] = this.getDefaultMonth();
    this.selectedMonth = defaultMonth;
    this.selectedYear = defaultYear;

    // Year range: current year -1 to +1
    const currentYear = new Date().getFullYear();
    for ( let y = currentYear - 1; y <= currentYear + 1; y++ ) {
      this.years.push( y );
    }
  }


  /**
   * Load data on init and subscribe to real-time folder changes.
   */
  ngOnInit(): void {
    // Debounce reload requests so rapid socket events don't spam the server
    this.reload$.pipe( debounceTime( 500 ), takeUntil( this.destroy$ ) )
      .subscribe( () => this.doLoadMonth() );

    this.loadMonth();
    this.loadTemplateInfo();

    this.socketService.on( 'folder-changes' ).pipe(
      debounceTime( 2000 ),
      takeUntil( this.destroy$ ),
    ).subscribe( () => this.loadMonth() );
  }


  /**
   * Clean up subscriptions.
   */
  ngOnDestroy(): void {
    this.destroy$.next();
    this.destroy$.complete();
  }


  /**
   * Queue a debounced reload of month data.
   */
  loadMonth(): void {
    this.reload$.next();
  }


  /**
   * Load song selections and folders for the selected month.
   * Immediately builds skeleton weeks from date math, then fills in data as responses arrive.
   */
  private doLoadMonth(): void {
    this.editingPastWeeks.clear();
    this.loading = true;

    // Immediately build skeleton weeks so cards render right away
    const sundays = this.computeSundays();
    this.weeks = sundays.map( date => ({
      date,
      label: this.formatDateLabel( date ),
      selection: null,
      folder: this.findFolder( date ),
    }));

    this.loadVerse();
    this.loadClosingSong();

    // Load week selections — merge into existing skeleton weeks
    this.http.get<Record<string, WeekSelection>>(
      `/api/week-song-selections?year=${ this.selectedYear }&month=${ this.selectedMonth }`
    ).subscribe({
      next: ( data ) => {
        const weekIndices = Object.keys( data ).map( Number ).sort();
        for ( const idx of weekIndices ) {
          if ( this.weeks[ idx ] ) {
            this.weeks[ idx ].selection = data[ String( idx ) ];
          }
        }
        this.loading = false;
      },
      error: () => { this.loading = false; },
    });

    // Load folders for thumbnails and status — merge into existing skeleton weeks
    this.http.get<FolderInfo[]>( '/api/folders' ).subscribe({
      next: ( data ) => {
        this.folders = data;
        this.weeks.forEach( w => {
          w.folder = this.findFolder( w.date );
        });
      },
      error: () => {},
    });
  }


  /**
   * Open the song selection dialog for a specific slot.
   *
   * @param weekIdx - index of the week
   * @param slot - song slot (1, 2, or 3)
   */
  selectSong( weekIdx: number, slot: number ): void {
    const week = this.weeks[ weekIdx ];
    const currentSong = week.selection?.songs?.[ slot - 1 ] || null;

    const dialogRef = this.dialog.open( SongSelectDialogComponent, {
      width: '900px',
      maxHeight: '80vh',
      data: { currentSong, slot },
    });

    dialogRef.afterClosed().subscribe( ( song: Song | undefined ) => {
      if ( !song ) return;

      this.http.post( '/api/update-song', {
        date: week.date,
        slot,
        song,
        weekIdx,
      }).subscribe({
        next: () => this.loadMonth(),
        error: ( err ) => console.error( 'Failed to update song:', err ),
      });
    });
  }


  /**
   * Get the YYYYMMDD date string for a week index.
   *
   * @param idx - zero-based week index
   * @returns YYYYMMDD formatted date
   */
  private getSundayDate( idx: number ): string {
    const sundays = this.computeSundays();
    return sundays[ idx ] || '';
  }


  /**
   * Compute YYYYMMDD strings for all Sundays in the selected month.
   *
   * @returns array of YYYYMMDD strings
   */
  private computeSundays(): string[] {
    const sundays: string[] = [];
    const daysInMonth = new Date( this.selectedYear, this.selectedMonth, 0 ).getDate();
    for ( let day = 1; day <= daysInMonth; day++ ) {
      const d = new Date( this.selectedYear, this.selectedMonth - 1, day );
      if ( d.getDay() === 0 ) {
        const mm = String( this.selectedMonth ).padStart( 2, '0' );
        const dd = String( day ).padStart( 2, '0' );
        sundays.push( `${ this.selectedYear }${ mm }${ dd }` );
      }
    }
    return sundays;
  }


  /**
   * Format a week index into a readable Sunday label.
   *
   * @param idx - zero-based week index
   * @returns formatted label like "Sun, Feb 2"
   */
  private formatSundayLabel( idx: number ): string {
    const dateStr = this.getSundayDate( idx );
    return this.formatDateLabel( dateStr );
  }


  /**
   * Format a YYYYMMDD string into a readable date label.
   *
   * @param dateStr - YYYYMMDD format
   * @returns formatted label like "Sunday, Feb 2"
   */
  formatDateLabel( dateStr: string ): string {
    if ( !dateStr || dateStr.length !== 8 ) return dateStr;
    const y = parseInt( dateStr.slice( 0, 4 ) );
    const m = parseInt( dateStr.slice( 4, 6 ) ) - 1;
    const d = parseInt( dateStr.slice( 6, 8 ) );
    const date = new Date( y, m, d );
    return date.toLocaleDateString( 'en-US', { weekday: 'long', month: 'short', day: 'numeric' });
  }


  /**
   * Find folder info for a given YYYYMMDD date.
   *
   * @param date - YYYYMMDD string
   * @returns matching folder or null
   */
  private findFolder( date: string ): FolderInfo | null {
    return this.folders.find( f => f.name === date ) || null;
  }


  /**
   * Get display text for a song in a slot.
   *
   * @param song - song object or null
   * @returns display string
   */
  songDisplay( song: Song | null ): string {
    if ( !song ) return 'Empty';
    if ( song.number ) return `${ song.number } - ${ song.name }`;
    return song.name;
  }


  /**
   * Check if a YYYYMMDD date is in the past (before today).
   *
   * @param dateStr - YYYYMMDD format
   * @returns true if the date is before today
   */
  isPastSunday( dateStr: string ): boolean {
    if ( !dateStr || dateStr.length !== 8 ) return false;
    const y = parseInt( dateStr.slice( 0, 4 ) );
    const m = parseInt( dateStr.slice( 4, 6 ) ) - 1;
    const d = parseInt( dateStr.slice( 6, 8 ) );
    const sundayDate = new Date( y, m, d );
    const today = new Date();
    today.setHours( 0, 0, 0, 0 );
    return sundayDate < today;
  }


  /**
   * Remove a chorus from a week and persist the change.
   *
   * @param weekIdx - index of the week
   * @param chorusIdx - index of the chorus to remove
   */
  removeChorus( weekIdx: number, chorusIdx: number ): void {
    const week = this.weeks[ weekIdx ];
    if ( !week.selection?.choruses ) return;

    // Remove the chorus from the local array
    const updatedChoruses = [ ...week.selection.choruses ];
    updatedChoruses.splice( chorusIdx, 1 );

    // Persist the full chorus list
    this.http.post( '/api/chorus-update', {
      weekIdx,
      slot: 'choruses',
      choruses: updatedChoruses,
      date: week.date,
    }).subscribe({
      next: () => this.loadMonth(),
      error: ( err ) => console.error( 'Failed to remove chorus:', err ),
    });
  }


  /**
   * Open the chorus edit dialog to change or remove an existing chorus.
   *
   * @param weekIdx - index of the week
   * @param chorusIdx - index of the chorus to edit
   */
  editChorus( weekIdx: number, chorusIdx: number ): void {
    const week = this.weeks[ weekIdx ];
    const chorus = week.selection?.choruses?.[ chorusIdx ];
    if ( !chorus ) return;

    const dialogRef = this.dialog.open( ChorusEditDialogComponent, {
      width: '400px',
      data: { chorus, chorusIndex: chorusIdx },
    });

    dialogRef.afterClosed().subscribe( ( result: ChorusEditDialogResult | undefined ) => {
      if ( !result ) return;

      if ( result.action === 'remove' ) {
        this.removeChorus( weekIdx, chorusIdx );
      } else if ( result.action === 'change' && result.song ) {
        const updatedChoruses = [ ...( week.selection?.choruses || [] ) ];
        updatedChoruses[ chorusIdx ] = result.song;

        this.http.post( '/api/chorus-update', {
          weekIdx,
          slot: 'choruses',
          choruses: updatedChoruses,
          date: week.date,
        }).subscribe({
          next: () => this.loadMonth(),
          error: ( err ) => console.error( 'Failed to update chorus:', err ),
        });
      }
    });
  }


  /**
   * Open the song selection dialog to add a chorus.
   *
   * @param weekIdx - index of the week
   */
  addChorus( weekIdx: number ): void {
    const dialogRef = this.dialog.open( SongSelectDialogComponent, {
      width: '900px',
      maxHeight: '80vh',
      data: { currentSong: null, slot: 'chorus' },
    });

    dialogRef.afterClosed().subscribe( ( song: Song | undefined ) => {
      if ( !song ) return;

      const week = this.weeks[ weekIdx ];
      const currentChoruses = week.selection?.choruses || [];
      const updatedChoruses = [ ...currentChoruses, song ];

      // Persist the full chorus list
      this.http.post( '/api/chorus-update', {
        weekIdx,
        slot: 'choruses',
        choruses: updatedChoruses,
        date: week.date,
      }).subscribe({
        next: () => this.loadMonth(),
        error: ( err ) => console.error( 'Failed to add chorus:', err ),
      });
    });
  }


  /**
   * Open the notes editor dialog for a week.
   *
   * @param week - the week object containing folder info
   */
  openNotes( week: { date: string; folder: FolderInfo | null } ): void {
    if ( !week.date ) return;

    const dialogRef = this.dialog.open( NotesDialogComponent, {
      width: '900px',
      data: { folderName: week.date },
    });

    dialogRef.afterClosed().subscribe( () => this.loadMonth() );
  }


  /**
   * Load the bible verse reference and text for the selected month/year.
   */
  loadVerse(): void {
    this.http.get<{ reference: string; text: string }>(
      `/api/verse?year=${ this.selectedYear }&month=${ this.selectedMonth }`
    ).subscribe({
      next: ( data ) => {
        this.verseReference = data.reference || '';
        this.verseText = data.text || '';
      },
      error: () => {
        this.verseReference = '';
        this.verseText = '';
      },
    });
  }


  /**
   * Load template file info for the footer.
   */
  private loadTemplateInfo(): void {
    this.http.get<{ name: string; lastModified: string; size: string }>( '/api/template-info' ).subscribe({
      next: ( data ) => { this.templateInfo = data; },
      error: () => {},
    });
  }


  /**
   * Open the bible verse editor dialog.
   */
  openVerseDialog(): void {
    const dialogRef = this.dialog.open( VerseDialogComponent, {
      width: '550px',
      data: {
        year: this.selectedYear,
        month: this.selectedMonth,
        monthName: this.months[ this.selectedMonth - 1 ],
        reference: this.verseReference,
        text: this.verseText,
      } as VerseDialogData,
    });

    dialogRef.afterClosed().subscribe( ( result: VerseDialogResult | undefined ) => {
      if ( result ) {
        this.verseReference = result.reference;
        this.verseText = result.text;
      }
    });
  }


  /**
   * Load the closing song for the selected month/year.
   */
  loadClosingSong(): void {
    this.http.get<{ song: { name: string; number?: string } | null }>(
      `/api/closing-song?year=${ this.selectedYear }&month=${ this.selectedMonth }`
    ).subscribe({
      next: ( data ) => {
        this.closingSong = data.song || null;
      },
      error: () => {
        this.closingSong = null;
      },
    });
  }


  /**
   * Open the song selection dialog to pick a closing song for the month.
   */
  selectClosingSong(): void {
    const dialogRef = this.dialog.open( SongSelectDialogComponent, {
      width: '900px',
      maxHeight: '80vh',
      data: { currentSong: this.closingSong, slot: 'closing' },
    });

    dialogRef.afterClosed().subscribe( ( song: Song | undefined ) => {
      if ( song === undefined ) return;

      this.http.post( '/api/closing-song', {
        year: this.selectedYear,
        month: this.selectedMonth,
        song,
      }).subscribe({
        next: () => this.loadClosingSong(),
        error: ( err ) => console.error( 'Failed to save closing song:', err ),
      });
    });
  }


  /**
   * Approve the presentation for a week, removing "TODO" from the filename.
   *
   * @param week - the week whose presentation to approve
   */
  approvePresentation( week: { date: string; folder: FolderInfo | null } ): void {
    if ( !week.date || !week.folder?.hasPre || week.folder?.isApproved ) return;

    this.http.put( `/api/folders/${ week.date }/approve`, {} ).subscribe({
      next: () => this.loadMonth(),
      error: ( err ) => console.error( 'Failed to approve presentation:', err ),
    });
  }


  /**
   * Open the YouTube URL dialog for a week.
   *
   * @param week - the week to set the YouTube URL for
   */
  openYoutubeDialog( week: { date: string; folder: FolderInfo | null } ): void {
    const ref = this.dialog.open( YoutubeDialogComponent, {
      data: {
        folder: week.date,
        date: week.date,
        currentUrl: week.folder?.youtubeUrl || null,
      } as YoutubeDialogData,
      width: '480px',
    } );

    ref.afterClosed().subscribe( ( result: { url: string | null } | undefined ) => {
      if ( result !== undefined ) {
        this.loadMonth();
      }
    } );
  }


  /**
   * Open the folder creation dialog.
   */
  runFolders(): void {
    this.dialog.open( ConfirmDialogComponent, {
      data: {
        title: 'Create Folders for Sundays',
        apiEndpoint: '/api/run-folders',
        month: this.selectedMonth,
        monthName: this.months[ this.selectedMonth - 1 ],
        year: this.selectedYear,
      },
    }).afterClosed().subscribe( () => this.loadMonth() );
  }


  /**
   * Open the template copy dialog.
   */
  copyTemplate(): void {
    this.dialog.open( ConfirmDialogComponent, {
      data: {
        title: 'Copy Template for Sundays',
        apiEndpoint: '/api/copy-template',
        month: this.selectedMonth,
        monthName: this.months[ this.selectedMonth - 1 ],
        year: this.selectedYear,
        showOverwrite: this.anyHasPre,
      },
    }).afterClosed().subscribe( () => this.loadMonth() );
  }


  /**
   * True when all Sundays for the month already have folders.
   */
  get allFoldersExist(): boolean {
    if ( this.weeks.length === 0 ) return false;
    return this.weeks.every( w => w.folder !== null );
  }


  /**
   * Number of days until the first Sunday of the selected month, or null if already past.
   */
  get daysUntilFirstSunday(): number | null {
    if ( this.weeks.length === 0 ) return null;
    const dateStr = this.weeks[ 0 ].date;
    const year = parseInt( dateStr.slice( 0, 4 ) );
    const month = parseInt( dateStr.slice( 4, 6 ) ) - 1;
    const day = parseInt( dateStr.slice( 6, 8 ) );
    const firstSunday = new Date( year, month, day );
    const now = new Date();
    const today = new Date( now.getFullYear(), now.getMonth(), now.getDate() );
    const diff = Math.ceil( ( firstSunday.getTime() - today.getTime() ) / ( 1000 * 60 * 60 * 24 ) );
    return diff > 0 ? diff : null;
  }


  /**
   * True when any Sunday folder already has a presentation file.
   */
  get anyHasPre(): boolean {
    return this.weeks.some( w => w.folder?.hasPre );
  }


  /**
   * True when the selected month is entirely in the past (last day of month < today).
   */
  get isPreviousMonth(): boolean {
    const now = new Date();
    const lastDay = new Date( this.selectedYear, this.selectedMonth, 0 );
    const today = new Date( now.getFullYear(), now.getMonth(), now.getDate() );
    return lastDay < today;
  }


  /**
   * Whether a week card is locked (past and not being edited).
   *
   * @param week - the week object
   * @returns true if the card should be in read-only mode
   */
  isWeekLocked( week: { date: string } ): boolean {
    return this.isPastSunday( week.date ) && !this.editingPastWeeks.has( week.date );
  }


  /**
   * Toggle edit mode for a past week.
   *
   * @param week - the week to toggle
   */
  toggleEditPastWeek( week: { date: string } ): void {
    if ( this.editingPastWeeks.has( week.date ) ) {
      this.editingPastWeeks.delete( week.date );
    } else {
      this.editingPastWeeks.add( week.date );
    }
  }


  /**
   * TrackBy function for week cards to prevent unnecessary DOM recreation.
   *
   * @param index - array index
   * @param week - week object
   * @returns the week's date string as a stable identity
   */
  trackByDate( index: number, week: { date: string } ): string {
    return week.date;
  }


  /**
   * Get the default month/year based on current date.
   * If after the 15th of the month, defaults to next month.
   *
   * @returns tuple of [month (1-12), year]
   */
  private getDefaultMonth(): [ number, number ] {
    const now = new Date();
    let month = now.getMonth() + 1;
    let year = now.getFullYear();
    const day = now.getDate();

    // If after the 15th, default to next month
    if ( day > 15 ) {
      month++;
      if ( month > 12 ) {
        month = 1;
        year++;
      }
    }

    return [ month, year ];
  }
}
