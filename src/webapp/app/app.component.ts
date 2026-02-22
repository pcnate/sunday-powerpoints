import { Component } from '@angular/core';
import { MatTabsModule } from '@angular/material/tabs';
import { MatIconModule } from '@angular/material/icon';
import { UpcomingComponent } from './upcoming/upcoming.component';
import { LibraryComponent } from './library/library.component';
import { OutstandingComponent } from './outstanding/outstanding.component';
import { JobsComponent } from './jobs/jobs.component';


/**
 * Root application component with tabbed navigation.
 */
@Component({
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
})
export class AppComponent {}
