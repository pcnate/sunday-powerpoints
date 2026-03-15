import { Injectable, OnDestroy } from '@angular/core';
import { BehaviorSubject, Subject } from 'rxjs';


/**
 * Status snapshot pushed from the Rust worker on connect and lag recovery.
 */
export interface WorkerStatus {
  phase: string;
  connected: boolean;
  current_job_id: number | null;
  current_job_type: string | null;
  jobs_completed: number;
  jobs_failed: number;
  uptime_secs: number;
  worker_id: string;
  server_url: string;
  schedule_enabled: boolean;
  shift_start: string;
  shift_end: string;
  force_on_shift: boolean;
  paused_types: string[];
  active_types: string[];
}


/**
 * A typed WebSocket event from the Rust worker.
 */
export interface WorkerEvent {
  type: string;
  data?: Record<string, unknown>;
}


/**
 * WebSocket service for real-time push updates from the Rust worker.
 *
 * Maintains a persistent connection with auto-reconnect. Exposes
 * a `status$` BehaviorSubject that always reflects the latest
 * worker state, and event subjects for queue/config changes.
 */
@Injectable({ providedIn: 'root' })
export class WorkerWsService implements OnDestroy {

  /** Current worker status, updated incrementally via WebSocket events. */
  readonly status$ = new BehaviorSubject<WorkerStatus | null>( null );

  /** Whether the WebSocket itself is connected to the worker. */
  readonly wsConnected$ = new BehaviorSubject<boolean>( false );

  /** Emits when the job queue should be refreshed. */
  readonly queueStale$ = new Subject<void>();

  /** Emits when config was reloaded — UI should re-fetch GET /api/config. */
  readonly configReloaded$ = new Subject<void>();

  /** Raw event stream for components that need fine-grained control. */
  readonly events$ = new Subject<WorkerEvent>();

  private socket: WebSocket | null = null;
  private reconnectTimer: ReturnType<typeof setTimeout> | null = null;
  private reconnectDelay = 1000;
  private destroyed = false;


  constructor() {
    this.connect();
  }


  ngOnDestroy(): void {
    this.destroyed = true;
    this.clearReconnect();
    this.socket?.close();
  }


  /**
   * Establish the WebSocket connection.
   */
  connect(): void {
    if ( this.destroyed ) return;

    const protocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
    const url = `${ protocol }//${ location.host }/ws`;

    this.socket = new WebSocket( url );

    this.socket.onopen = () => {
      this.wsConnected$.next( true );
      this.reconnectDelay = 1000;
    };

    this.socket.onclose = () => {
      this.wsConnected$.next( false );
      this.scheduleReconnect();
    };

    this.socket.onerror = () => {
      this.socket?.close();
    };

    this.socket.onmessage = ( msg ) => {
      try {
        const event: WorkerEvent = JSON.parse( msg.data );
        this.handleEvent( event );
      } catch {
        // ignore malformed messages
      }
    };
  }


  /**
   * Process an incoming WebSocket event and update local state.
   *
   * @param event - parsed WebSocket event
   */
  private handleEvent( event: WorkerEvent ): void {
    this.events$.next( event );
    const status = this.status$.value;

    switch ( event.type ) {
      case 'status_snapshot': {
        this.status$.next( event.data as unknown as WorkerStatus );
        break;
      }
      case 'phase_change': {
        if ( status ) {
          this.status$.next({ ...status, phase: event.data![ 'phase' ] as string });
        }
        break;
      }
      case 'connection': {
        if ( status ) {
          this.status$.next({ ...status, connected: event.data![ 'connected' ] as boolean });
        }
        break;
      }
      case 'job_claimed': {
        if ( status ) {
          this.status$.next({
            ...status,
            current_job_id: event.data![ 'job_id' ] as number,
            current_job_type: event.data![ 'job_type' ] as string,
          });
        }
        break;
      }
      case 'job_completed':
      case 'job_failed': {
        this.queueStale$.next();
        break;
      }
      case 'job_cleared': {
        if ( status ) {
          this.status$.next({
            ...status,
            current_job_id: null,
            current_job_type: null,
          });
        }
        break;
      }
      case 'stats_update': {
        if ( status ) {
          this.status$.next({
            ...status,
            jobs_completed: event.data![ 'jobs_completed' ] as number,
            jobs_failed: event.data![ 'jobs_failed' ] as number,
          });
        }
        break;
      }
      case 'force_on_shift': {
        if ( status ) {
          this.status$.next({ ...status, force_on_shift: event.data![ 'enabled' ] as boolean });
        }
        break;
      }
      case 'type_pause_change': {
        if ( status ) {
          const jobType = event.data![ 'job_type' ] as string;
          const isPaused = event.data![ 'paused' ] as boolean;
          const pausedTypes = status.paused_types.filter( t => t !== jobType );
          if ( isPaused ) pausedTypes.push( jobType );
          this.status$.next({ ...status, paused_types: pausedTypes });
        }
        break;
      }
      case 'config_reloaded': {
        this.configReloaded$.next();
        break;
      }
    }
  }


  /**
   * Schedule a reconnect attempt with exponential backoff.
   */
  private scheduleReconnect(): void {
    if ( this.destroyed ) return;
    this.clearReconnect();

    this.reconnectTimer = setTimeout( () => {
      this.connect();
    }, this.reconnectDelay );

    this.reconnectDelay = Math.min( this.reconnectDelay * 2, 30000 );
  }


  /**
   * Clear any pending reconnect timer.
   */
  private clearReconnect(): void {
    if ( this.reconnectTimer ) {
      clearTimeout( this.reconnectTimer );
      this.reconnectTimer = null;
    }
  }
}
