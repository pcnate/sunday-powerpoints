import { Injectable, OnDestroy } from '@angular/core';
import { io, Socket } from 'socket.io-client';
import { Observable, BehaviorSubject, Subject } from 'rxjs';

const SOCKET_URL = '/';


/**
 * SocketService connects to the backend socket.io server,
 * exposing socket events as RxJS observables and caching data for consumers.
 *
 * Usage:
 *   - Inject SocketService in any component or service.
 *   - Use on<T>( event ) to subscribe to socket events as observables.
 *   - Use emit( event, data ) to send events to the server.
 */
@Injectable({ providedIn: 'root' })
export class SocketService implements OnDestroy {
  private socket: Socket;
  private folders$ = new BehaviorSubject<any[]>([]);
  private songs$ = new BehaviorSubject<any[]>([]);
  private destroyed$ = new Subject<void>();

  constructor() {
    this.socket = io( SOCKET_URL, {
      transports: [ 'websocket' ],
      autoConnect: true,
    });

    this.socket.on( 'folder-changes', ( folders: any[] ) => {
      this.folders$.next( folders );
    });

    this.socket.on( 'song-changes', ( songs: any[] ) => {
      this.songs$.next( songs );
    });
  }


  /**
   * Subscribe to a socket.io event as an RxJS Observable.
   *
   * @param event - socket event name
   * @returns observable that emits event data
   */
  on<T = any>( event: string ): Observable<T> {
    return new Observable<T>( subscriber => {
      const handler = ( data: T ) => subscriber.next( data );
      this.socket.on( event, handler as any );
      return () => this.socket.off( event, handler as any );
    });
  }


  /**
   * Emit an event to the socket.io server.
   *
   * @param event - socket event name
   * @param data - optional payload
   */
  emit( event: string, data?: any ): void {
    this.socket.emit( event, data );
  }


  /**
   * Get cached folders as observable (auto-updates on folder-changes).
   */
  getFolders(): Observable<any[]> {
    return this.folders$.asObservable();
  }


  /**
   * Get cached songs as observable (auto-updates on song-changes).
   */
  getSongs(): Observable<any[]> {
    return this.songs$.asObservable();
  }


  /**
   * Disconnect and cleanup on destroy.
   */
  ngOnDestroy(): void {
    this.destroyed$.next();
    this.destroyed$.complete();
    this.socket.disconnect();
  }
}
