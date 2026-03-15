import { Component, OnInit, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { RouterLink } from '@angular/router';
import { HttpClient } from '@angular/common/http';
import { FormsModule } from '@angular/forms';
import { MatTableModule } from '@angular/material/table';
import { MatSortModule, Sort } from '@angular/material/sort';
import { MatButtonModule } from '@angular/material/button';
import { MatSelectModule } from '@angular/material/select';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatIconModule } from '@angular/material/icon';
import { MatChipsModule } from '@angular/material/chips';
import { MatTooltipModule } from '@angular/material/tooltip';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { MatCardModule } from '@angular/material/card';
import { MatDialog, MatDialogModule } from '@angular/material/dialog';
import { Subject, takeUntil, interval, merge, debounceTime } from 'rxjs';
import { SocketService } from '../../socket.service';
import { AdminComponent } from '../admin/admin.component';


/**
 * A job from the API.
 */
interface Job {
  id: number;
  type: string;
  status: string;
  priority: number;
  sunday_date: string;
  input_path: string;
  output_path: string | null;
  metadata: Record<string, unknown> | null;
  error_message: string | null;
  worker_id: string | null;
  retry_count: number;
  max_retries: number;
  created_at: string;
  started_at: string | null;
  heartbeat_at: string | null;
}


/**
 * A worker from the tracking API.
 */
interface WorkerInfo {
  id: string;
  types: string[];
  current_job_id: number | null;
  current_job_type: string | null;
  last_seen: string;
  first_seen: string;
  jobs_completed: number;
  jobs_failed: number;
  status: 'idle' | 'executing' | 'stale';
}


/**
 * Jobs component displaying the job queue and connected workers.
 */
@Component({
  selector: 'app-jobs',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatTableModule,
    MatSortModule,
    MatButtonModule,
    MatSelectModule,
    MatFormFieldModule,
    MatIconModule,
    MatChipsModule,
    MatTooltipModule,
    MatProgressSpinnerModule,
    MatCardModule,
    MatDialogModule,
    RouterLink,
  ],
  templateUrl: './jobs.component.html',
  styleUrls: [ './jobs.component.scss' ]
})
export class JobsComponent implements OnInit, OnDestroy {

  jobs: Job[] = [];
  workers: WorkerInfo[] = [];
  totalJobs = 0;
  loading = false;

  statusFilters: string[] = [ 'pending', 'processing' ];
  typeFilters: string[] = [];

  displayedColumns = [ 'id', 'type', 'sunday_date', 'status', 'worker_id', 'retry', 'created_at', 'actions' ];

  private destroy$ = new Subject<void>();


  constructor(
    private http: HttpClient,
    private socketService: SocketService,
    private dialog: MatDialog
  ) {}


  /**
   * Initialize data loading and real-time subscriptions.
   */
  ngOnInit(): void {
    this.loadJobs();
    this.loadWorkers();

    // Refresh jobs on socket events (debounced to avoid request storms)
    merge(
      this.socketService.on( 'job:created' ),
      this.socketService.on( 'job:updated' ),
    ).pipe( debounceTime( 500 ), takeUntil( this.destroy$ ) )
      .subscribe( () => this.loadJobs() );

    // Periodic refresh for workers (they update via heartbeat, no socket event)
    interval( 10000 ).pipe( takeUntil( this.destroy$ ) )
      .subscribe( () => this.loadWorkers() );
  }


  /**
   * Clean up subscriptions on destroy.
   */
  ngOnDestroy(): void {
    this.destroy$.next();
    this.destroy$.complete();
  }


  /**
   * Load jobs from the API with current filters.
   */
  loadJobs(): void {
    let url = `/api/jobs?limit=50`;
    if ( this.statusFilters.length > 0 ) url += `&status=${ this.statusFilters.join( ',' ) }`;
    if ( this.typeFilters.length > 0 ) url += `&type=${ this.typeFilters.join( ',' ) }`;

    this.http.get<{ jobs: Job[]; total: number }>( url )
      .subscribe({
        next: ( data ) => {
          this.jobs = data.jobs;
          this.totalJobs = data.total;
        },
        error: ( err ) => console.error( 'Failed to load jobs:', err ),
      });
  }


  /**
   * Load connected workers from the API.
   */
  loadWorkers(): void {
    this.http.get<{ workers: WorkerInfo[] }>( '/api/jobs/workers' )
      .subscribe({
        next: ( data ) => { this.workers = data.workers; },
        error: ( err ) => console.error( 'Failed to load workers:', err ),
      });
  }


  /**
   * Cancel a job by ID.
   *
   * @param job - the job to cancel
   */
  cancelJob( job: Job ): void {
    if ( !confirm( `Cancel job #${ job.id }?` ) ) return;

    this.http.put( `/api/jobs/${ job.id }`, { status: 'cancelled' } )
      .subscribe({
        next: () => this.loadJobs(),
        error: ( err ) => console.error( 'Failed to cancel job:', err ),
      });
  }


  /**
   * Sort the jobs table by column.
   *
   * @param sort - the sort event from MatSort
   */
  sortJobs( sort: Sort ): void {
    if ( !sort.active || sort.direction === '' ) {
      return;
    }

    this.jobs = [ ...this.jobs ].sort( ( a, b ) => {
      const asc = sort.direction === 'asc' ? 1 : -1;
      switch ( sort.active ) {
        case 'id': return ( a.id - b.id ) * asc;
        case 'type': return a.type.localeCompare( b.type ) * asc;
        case 'sunday_date': return a.sunday_date.localeCompare( b.sunday_date ) * asc;
        case 'status': return a.status.localeCompare( b.status ) * asc;
        case 'worker_id': return ( a.worker_id || '' ).localeCompare( b.worker_id || '' ) * asc;
        case 'retry': return ( a.retry_count - b.retry_count ) * asc;
        case 'created_at': return ( a.created_at || '' ).localeCompare( b.created_at || '' ) * asc;
        default: return 0;
      }
    });
  }


  /**
   * Get a CSS class for a job status badge.
   *
   * @param status - the job status string
   * @returns CSS class name
   */
  statusClass( status: string ): string {
    return `badge-${ status }`;
  }


  /**
   * Get a CSS class for a worker status indicator.
   *
   * @param status - the worker status
   * @returns CSS class name
   */
  workerStatusClass( status: string ): string {
    switch ( status ) {
      case 'executing': return 'dot-yellow';
      case 'idle': return 'dot-green';
      case 'stale': return 'dot-red';
      default: return 'dot-gray';
    }
  }


  /**
   * Build a routerLink path to the planning view for a sunday_date's month.
   *
   * @param sundayDate - YYYYMMDD format date string
   * @returns route path segments
   */
  monthRoute( sundayDate: string ): string[] {
    const year = sundayDate.substring( 0, 4 );
    const month = String( parseInt( sundayDate.substring( 4, 6 ), 10 ) );
    return [ '/planning', year, month ];
  }


  /**
   * Format a date string for display.
   *
   * @param dateStr - ISO date string
   * @returns formatted relative or short date
   */
  formatDate( dateStr: string | null ): string {
    if ( !dateStr ) return '-';
    const d = new Date( dateStr );
    const now = new Date();
    const diffMs = now.getTime() - d.getTime();
    const diffMins = Math.floor( diffMs / 60000 );

    if ( diffMins < 1 ) return 'just now';
    if ( diffMins < 60 ) return `${ diffMins }m ago`;
    if ( diffMins < 1440 ) return `${ Math.floor( diffMins / 60 ) }h ago`;
    return d.toLocaleDateString();
  }


  /**
   * Open the job settings dialog.
   */
  openSettings(): void {
    this.dialog.open( AdminComponent, {
      width: '650px',
    });
  }
}
