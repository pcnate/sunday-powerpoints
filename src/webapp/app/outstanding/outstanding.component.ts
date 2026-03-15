import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterLink } from '@angular/router';
import { HttpClient } from '@angular/common/http';
import { MatIconModule } from '@angular/material/icon';
import { MatButtonModule } from '@angular/material/button';
import { MatButtonToggleModule } from '@angular/material/button-toggle';
import { MatProgressBarModule } from '@angular/material/progress-bar';
import { MatTooltipModule } from '@angular/material/tooltip';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { MatDialogModule, MatDialog } from '@angular/material/dialog';
import { Subject, takeUntil } from 'rxjs';
import { SocketService } from '../../socket.service';
import { NotesDialogComponent } from '../shared/notes-dialog.component';
import { KdenliveDialogComponent } from '../shared/kdenlive-dialog.component';
import { QueueJobDialogComponent } from '../shared/queue-job-dialog.component';
import { YoutubeDialogComponent } from '../upcoming/youtube-dialog.component';
import { SimpleConfirmDialogComponent } from '../shared/simple-confirm-dialog.component';


/**
 * Raw folder metadata from the API, including pipeline fields.
 */
interface RawFolder {
  name: string;
  hasPre: boolean;
  isApproved: boolean;
  hasMp4: boolean;
  hasVideo: boolean;
  hasThumbnail: boolean;
  thumbnail: string | null;
  videoFileName: string | null;
  hasSong1: boolean;
  hasSong2: boolean;
  hasSong3: boolean;
  songSlots?: string[];
  hasNotes: boolean;
  hasKdenlive: boolean;
  hasProduction: boolean;
  hasTranscription: boolean;
  hasSummary: boolean;
  youtubeUrl: string | null;
  backlog: boolean;
  archived: boolean;
  relativePath: string;
}


/**
 * Extended folder info with computed stage-aware completion fields.
 *
 * Stage 1 (pre-service): songs, presentation, approval
 * Stage 2 (post-service): video → kdenlive → production → transcription → sermon → youtube
 * Only tracked after the Sunday has passed AND Stage 1 is complete.
 */
interface FolderInfo extends RawFolder {
  completionCount: number;
  totalItems: number;
  missingItems: string[];
  isComplete: boolean;
  displayDate: string;
  isPast: boolean;
  stage: 1 | 2;
}


/**
 * Outstanding tab — shows all non-archived folders with pipeline completion status.
 * Defaults to showing only folders that need attention.
 */
@Component({
  selector: 'app-outstanding',
  standalone: true,
  imports: [
    CommonModule,
    MatIconModule,
    MatButtonModule,
    MatButtonToggleModule,
    MatProgressBarModule,
    MatTooltipModule,
    MatProgressSpinnerModule,
    MatDialogModule,
    RouterLink,
  ],
  templateUrl: './outstanding.component.html',
  styleUrls: [ './outstanding.component.scss' ]
})
export class OutstandingComponent implements OnInit, OnDestroy {

  allFolders: FolderInfo[] = [];
  filteredFolders: FolderInfo[] = [];
  filterMode: 'attention' | 'all' = 'attention';
  loading = true;

  totalCount = 0;
  needsAttentionCount = 0;
  completeCount = 0;

  /** Maps sunday_date → job_type → status for active (pending/processing) jobs */
  activeJobs: Record<string, Record<string, string>> = {};

  private destroy$ = new Subject<void>();


  constructor(
    private http: HttpClient,
    private dialog: MatDialog,
    private socketService: SocketService
  ) {}


  /**
   * Load data on init and subscribe to real-time folder changes.
   */
  ngOnInit(): void {
    this.loadFolders();

    this.socketService.on( 'folder-changes' ).pipe( takeUntil( this.destroy$ ) )
      .subscribe( () => this.loadFolders() );
    this.socketService.on( 'job:created' ).pipe( takeUntil( this.destroy$ ) )
      .subscribe( () => this.loadActiveJobs() );
    this.socketService.on( 'job:updated' ).pipe( takeUntil( this.destroy$ ) )
      .subscribe( () => {
        this.loadActiveJobs();
        this.loadFolders();
      });
  }


  /**
   * Clean up subscriptions.
   */
  ngOnDestroy(): void {
    this.destroy$.next();
    this.destroy$.complete();
  }


  /**
   * Load all non-archived folders from API and enrich with computed fields.
   */
  loadFolders(): void {
    this.http.get<RawFolder[]>( '/api/folders' ).subscribe({
      next: ( data ) => {
        this.allFolders = this.enrichFolders( data );
        this.applyFilter();
        this.loading = false;
      },
      error: () => { this.loading = false; },
    });
    this.loadActiveJobs();
  }


  /**
   * Fetch pending and processing jobs to show queued/active status on icons.
   */
  loadActiveJobs(): void {
    this.http.get<{ jobs: { sunday_date: string; type: string; status: string }[] }>(
      '/api/jobs?status=pending,processing&type=transcode,transcription,claude-processing&limit=200'
    ).subscribe({
      next: ( data ) => {
        const map: Record<string, Record<string, string>> = {};
        for ( const job of data.jobs ) {
          if ( !map[ job.sunday_date ] ) map[ job.sunday_date ] = {};
          const existing = map[ job.sunday_date ][ job.type ];
          // processing takes priority over pending
          if ( !existing || job.status === 'processing' ) {
            map[ job.sunday_date ][ job.type ] = job.status;
          }
        }
        this.activeJobs = map;
      },
      error: () => { this.activeJobs = {}; },
    });
  }


  /**
   * Enrich raw folder data with stage-aware pipeline completion fields.
   *
   * Stage 1 (pre-service): songs (3) + presentation + approval — always relevant.
   * Stage 2 (post-service): video → kdenlive → production → transcription → summary → youtube.
   * Only counted after the Sunday has passed AND Stage 1 is complete.
   *
   * @param raw - array of raw folder objects from the API
   * @returns enriched folder array sorted newest first
   */
  enrichFolders( raw: RawFolder[] ): FolderInfo[] {
    const today = new Date();
    today.setHours( 0, 0, 0, 0 );

    return raw.filter( ( f ) => !f.archived ).map( ( f ) => {
      const isPast = this.parseDate( f.name ) < today;

      // Stage 1: pre-service checklist (dynamic song slots when available)
      const songChecks = ( f.songSlots && f.songSlots.length > 0 )
        ? f.songSlots.map( slotId => ({ key: `Song ${ slotId }`, has: true }) )
        : [
          { key: 'Song 1', has: f.hasSong1 },
          { key: 'Song 2', has: f.hasSong2 },
          { key: 'Song 3', has: f.hasSong3 },
        ];
      const stage1 = [
        ...songChecks,
        { key: 'Presentation', has: f.hasPre },
        { key: 'Approval', has: !f.hasPre || f.isApproved },
      ];

      // Stage 2: post-service pipeline (only for past dates with complete Stage 1)
      // If YouTube URL is set, the video pipeline is considered complete
      const ytDone = !!f.youtubeUrl;
      const stage2 = [
        { key: 'Video', has: f.hasVideo || ytDone },
        { key: 'Kdenlive', has: f.hasKdenlive || ytDone },
        { key: 'Production', has: f.hasProduction || ytDone },
        { key: 'Transcription', has: f.hasTranscription || ytDone },
        { key: 'Summary', has: f.hasSummary || ytDone },
        { key: 'YouTube', has: ytDone },
        { key: 'Archive', has: f.archived },
      ];

      const stage1Complete = stage1.every( ( c ) => c.has );
      const includeStage2 = ( isPast && stage1Complete ) || f.hasVideo || f.hasThumbnail;
      const checks = includeStage2 ? [ ...stage1, ...stage2 ] : stage1;
      const completionCount = checks.filter( ( c ) => c.has ).length;
      const totalItems = checks.length;
      const missingItems = checks.filter( ( c ) => !c.has ).map( ( c ) => c.key );

      return {
        ...f,
        completionCount,
        totalItems,
        missingItems,
        isComplete: completionCount === totalItems,
        displayDate: this.formatDate( f.name ),
        isPast,
        stage: includeStage2 ? 2 as const : 1 as const,
      };
    }).sort( ( a, b ) => b.name.localeCompare( a.name ) );
  }


  /**
   * Parse a YYYYMMDD string into a Date object.
   *
   * @param dateStr - YYYYMMDD format
   * @returns Date at midnight local time
   */
  parseDate( dateStr: string ): Date {
    const y = parseInt( dateStr.slice( 0, 4 ) );
    const m = parseInt( dateStr.slice( 4, 6 ) ) - 1;
    const d = parseInt( dateStr.slice( 6, 8 ) );
    return new Date( y, m, d );
  }


  /**
   * Apply the current filter mode and update summary stats.
   */
  applyFilter(): void {
    this.totalCount = this.allFolders.length;
    this.needsAttentionCount = this.allFolders.filter( ( f ) => !f.isComplete ).length;
    this.completeCount = this.totalCount - this.needsAttentionCount;

    if ( this.filterMode === 'attention' ) {
      this.filteredFolders = this.allFolders.filter( ( f ) => !f.isComplete );
    } else {
      this.filteredFolders = [ ...this.allFolders ];
    }
  }


  /**
   * Switch the filter mode and reapply.
   *
   * @param mode - 'attention' for incomplete only, 'all' for everything
   */
  setFilter( mode: 'attention' | 'all' ): void {
    this.filterMode = mode;
    this.applyFilter();
  }


  /**
   * Format a YYYYMMDD string into a readable date label.
   *
   * @param dateStr - YYYYMMDD format
   * @returns formatted label like "Sun, Feb 2, 2025"
   */
  formatDate( dateStr: string ): string {
    if ( !dateStr || dateStr.length !== 8 ) return dateStr;
    const y = parseInt( dateStr.slice( 0, 4 ) );
    const m = parseInt( dateStr.slice( 4, 6 ) ) - 1;
    const d = parseInt( dateStr.slice( 6, 8 ) );
    const date = new Date( y, m, d );
    return date.toLocaleDateString( 'en-US', {
      weekday: 'short',
      month: 'short',
      day: 'numeric',
      year: 'numeric',
    });
  }


  /**
   * Get the active job status for a folder and job type.
   *
   * @param folderName - YYYYMMDD folder name
   * @param jobType - job type (transcode, transcription, claude-processing)
   * @returns 'pending', 'processing', or null
   */
  jobStatus( folderName: string, jobType: string ): string | null {
    return this.activeJobs[ folderName ]?.[ jobType ] || null;
  }


  /**
   * Build a routerLink path to the planning view for a folder's month.
   *
   * @param name - YYYYMMDD folder name
   * @returns route path segments
   */
  monthRoute( name: string ): string[] {
    const year = name.substring( 0, 4 );
    const month = String( parseInt( name.substring( 4, 6 ), 10 ) );
    return [ '/planning', year, month ];
  }


  /**
   * Get completion percentage for a folder.
   *
   * @param folder - the folder info
   * @returns percentage 0-100
   */
  completionPercent( folder: FolderInfo ): number {
    return Math.round( ( folder.completionCount / folder.totalItems ) * 100 );
  }


  /**
   * Open the notes editor dialog for a folder.
   *
   * @param folder - the folder to edit notes for
   */
  openNotes( folder: FolderInfo ): void {
    this.dialog.open( NotesDialogComponent, {
      width: '600px',
      data: { folderName: folder.name },
    });
  }


  /**
   * Open the Kdenlive project creation dialog for a folder.
   *
   * @param folder - the folder to create a Kdenlive project for
   */
  openKdenlive( folder: FolderInfo ): void {
    this.dialog.open( KdenliveDialogComponent, {
      width: '550px',
      data: { folderName: folder.name },
    }).afterClosed().subscribe( ( created ) => {
      if ( created ) this.loadFolders();
    });
  }


  /**
   * Approve a Kdenlive project and queue a transcode job.
   *
   * @param folder - the folder with a Kdenlive project to approve
   */
  approveKdenlive( folder: FolderInfo ): void {
    this.http.post<{ ok: boolean; jobId: number }>( `/api/folders/${ folder.name }/approve-kdenlive`, {} ).subscribe({
      next: () => this.loadFolders(),
      error: ( err ) => console.error( 'Failed to approve Kdenlive:', err ),
    });
  }


  /**
   * Open the queue transcription dialog for a folder.
   *
   * @param folder - the folder to queue transcription for
   */
  queueTranscription( folder: FolderInfo ): void {
    this.dialog.open( QueueJobDialogComponent, {
      width: '550px',
      data: { folderName: folder.name, jobType: 'transcription' },
    }).afterClosed().subscribe( ( queued ) => {
      if ( queued ) this.loadFolders();
    });
  }


  /**
   * Open the queue claude-processing dialog for a folder.
   *
   * @param folder - the folder to queue claude processing for
   */
  queueClaude( folder: FolderInfo ): void {
    this.dialog.open( QueueJobDialogComponent, {
      width: '550px',
      data: { folderName: folder.name, jobType: 'claude-processing' },
    }).afterClosed().subscribe( ( queued ) => {
      if ( queued ) this.loadFolders();
    });
  }


  /**
   * Open the YouTube URL dialog to set or update the link.
   *
   * @param folder - the folder to set the YouTube URL for
   */
  /**
   * Archive a folder after confirming with a dialog.
   *
   * @param folder - the folder to archive
   */
  archiveFolder( folder: FolderInfo ): void {
    const year = folder.name.slice( 0, 4 );
    const lastSlash = folder.relativePath.lastIndexOf( '/' );
    const outputDir = folder.backlog
      ? folder.relativePath.substring( 0, folder.relativePath.lastIndexOf( '/', lastSlash - 1 ) )
      : folder.relativePath.substring( 0, lastSlash );
    const archivePath = `${ outputDir }/${ year }/${ folder.name }`;

    const dialogRef = this.dialog.open( SimpleConfirmDialogComponent, {
      width: '450px',
      data: {
        title: 'Archive Folder',
        message: `Archive ${ folder.displayDate }?`,
        detail: `${ folder.relativePath } → ${ archivePath }`,
        icon: 'archive',
        confirmLabel: 'Archive',
      },
    });

    dialogRef.afterClosed().subscribe( ( confirmed ) => {
      if ( !confirmed ) return;
      this.http.post<{ ok: boolean }>( `/api/folders/${ folder.name }/archive`, {} ).subscribe({
        next: () => this.loadFolders(),
        error: ( err ) => console.error( 'Failed to archive folder:', err ),
      });
    });
  }


  openYoutube( folder: FolderInfo ): void {
    this.dialog.open( YoutubeDialogComponent, {
      width: '500px',
      data: {
        folder: folder.name,
        date: folder.displayDate,
        currentUrl: folder.youtubeUrl,
      },
    }).afterClosed().subscribe( ( result ) => {
      if ( result !== undefined ) this.loadFolders();
    });
  }


}
