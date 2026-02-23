import { Component, OnInit } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Router } from '@angular/router';
import { Title } from '@angular/platform-browser';
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
export class AppComponent implements OnInit {

  selectedIndex = 0;
  private appName = 'App';


  constructor(
    private http: HttpClient,
    private titleService: Title,
    private router: Router
  ) {}


  /**
   * Fetch app name from server and sync initial tab from URL.
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
  }


  /**
   * Navigate to the route for the selected tab and update the title.
   *
   * @param index - the selected tab index
   */
  onTabChange( index: number ): void {
    this.selectedIndex = index;
    this.router.navigate( [ TABS[ index ].route ] );
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
