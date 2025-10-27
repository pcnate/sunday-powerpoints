import { Injectable, OnDestroy } from '@angular/core';
import io from 'socket.io-client';
import { Observable, BehaviorSubject, Subject } from 'rxjs';

// not an absolute path, adjust based on your environment
const SOCKET_URL = '/socket/';

/**
 * SocketService is responsible for connecting to the backend socket.io server,
 * exposing socket events as RxJS observables, and caching data for consumers.
 *
 * Usage:
 *   - Inject SocketService in any component or service.
 *   - Use on<T>(event: string) to subscribe to socket events as observables.
 *   - Use emit(event, data) to send events to the server.
 *   - Cached data (e.g., folders, songs) can be exposed as BehaviorSubjects.
 */
@Injectable({ providedIn: 'root' })
export class SocketService implements OnDestroy {
  private socket: SocketIOClient.Socket;
  // Example: cache for folders and songs
  private folders$ = new BehaviorSubject<any[]>([]);
  private songs$ = new BehaviorSubject<any[]>([]);
  // Used to clean up subscriptions on destroy
  private destroyed$ = new Subject<void>();

  constructor() {
    // Connect to the socket.io server (adjust URL as needed)
    this.socket = io( SOCKET_URL, {
      transports: ['websocket'],
      autoConnect: true,
    });

    // Listen for folder-changes and update cache
    this.socket.on('folder-changes', (folders: any[]) => {
      this.folders$.next(folders);
    });
    // Listen for song-changes and update cache
    this.socket.on('song-changes', (songs: any[]) => {
      this.songs$.next(songs);
    });
  }

  /**
   * Subscribe to a socket.io event as an RxJS Observable.
   * Usage: this.socketService.on<T>('eventName').subscribe(...)
   */
  on<T = any>(event: string): Observable<T> {
    return new Observable<T>(subscriber => {
      const handler = (data: T) => subscriber.next(data);
      this.socket.on(event, handler);
      // Cleanup on unsubscribe
      return () => this.socket.off(event, handler);
    });
  }

  /**
   * Emit an event to the socket.io server.
   */
  emit(event: string, data?: any) {
    this.socket.emit(event, data);
  }

  /**
   * Get cached folders as observable (auto-updates on folder-changes)
   */
  getFolders(): Observable<any[]> {
    return this.folders$.asObservable();
  }

  /**
   * Get cached songs as observable (auto-updates on song-changes)
   */
  getSongs(): Observable<any[]> {
    return this.songs$.asObservable();
  }

  /**
   * Disconnect and cleanup on destroy
   */
  ngOnDestroy() {
    this.destroyed$.next();
    this.destroyed$.complete();
    this.socket.disconnect();
  }
}
