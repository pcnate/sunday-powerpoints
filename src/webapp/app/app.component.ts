import { Component, OnInit, OnDestroy } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { NavigationEnd, Router } from '@angular/router';
import { Title } from '@angular/platform-browser';
import { Subject, filter, takeUntil } from 'rxjs';
import { MatTabsModule } from '@angular/material/tabs';
import { MatIconModule } from '@angular/material/icon';
import { UpcomingComponent } from './upcoming/upcoming.component';
import { LibraryComponent } from './library/library.component';
import { OutstandingComponent } from './outstanding/outstanding.component';
import { JobsComponent } from './jobs/jobs.component';
/** Tab definition mapping index to route and display name. */
const TABS = [
  { route: 'planning', label: 'Planning' },
  { route: 'library', label: 'Library' },
  { route: 'outstanding', label: 'Outstanding' },
  { route: 'jobs', label: 'Jobs' },
];


/**
 * Root application component with tabbed navigation tied to routes.
 */
@Component( {
  selector: 'app-root',
  standalone: true,
  imports: [
    MatTabsModule,
    MatIconModule,
    UpcomingComponent,
    LibraryComponent,
    OutstandingComponent,
    JobsComponent,
  ],
  templateUrl: './app.component.html',
  styleUrls: [ './app.component.scss' ]
} )
export class AppComponent implements OnInit, OnDestroy {

  selectedIndex = 0;
  private appName = 'App';
  private destroy$ = new Subject<void>();


  constructor(
    private http: HttpClient,
    private titleService: Title,
    private router: Router
  ) {}


  /**
   * Fetch app name from server and sync initial tab from URL.
   * Subscribe to router events so programmatic navigation (e.g. from dialogs)
   * also switches the active tab.
   */
  ngOnInit(): void {
    // Determine initial tab from current URL path
    const path = window.location.pathname.replace( /^\//, '' ).split( '/' )[ 0 ];
    const tabIndex = TABS.findIndex( t => t.route === path );
    this.selectedIndex = tabIndex >= 0 ? tabIndex : 0;

    this.http.get<{ appName: string }>( '/api/app-config' ).subscribe( {
      next: ( config ) => {
        this.appName = config.appName;
        this.updateTitle( this.selectedIndex );
      },
      error: () => this.updateTitle( this.selectedIndex ),
    } );

    // React to programmatic navigation (e.g. dialog → router.navigate)
    this.router.events.pipe(
      filter( ( e ): e is NavigationEnd => e instanceof NavigationEnd ),
      takeUntil( this.destroy$ ),
    ).subscribe( ( e ) => {
      const segment = e.urlAfterRedirects.replace( /^\//, '' ).split( '/' )[ 0 ];
      const idx = TABS.findIndex( t => t.route === segment );
      if ( idx >= 0 && idx !== this.selectedIndex ) {
        this.selectedIndex = idx;
        this.updateTitle( idx );
      }
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
   * Navigate to the route for the selected tab and update the title.
   * Skips navigation if the URL already belongs to the target tab (e.g. when
   * a dialog navigated to /planning/2025/3 — we don't want to clobber it
   * with a bare /planning).
   *
   * @param index - the selected tab index
   */
  onTabChange( index: number ): void {
    this.selectedIndex = index;
    const currentSegment = window.location.pathname.replace( /^\//, '' ).split( '/' )[ 0 ];
    if ( currentSegment !== TABS[ index ].route ) {
      this.router.navigate( [ TABS[ index ].route ] );
    }
    this.updateTitle( index );
  }


  /**
   * Set the browser tab title to "AppName - TabName".
   *
   * @param index - the tab index
   */
  private updateTitle( index: number ): void {
    const tab = TABS[ index ]?.label || 'App';
    this.titleService.setTitle( `${ this.appName } - ${ tab }` );
  }
}
