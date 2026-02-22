import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { HttpClient } from '@angular/common/http';
import { MatCardModule } from '@angular/material/card';
import { MatIconModule } from '@angular/material/icon';


/**
 * Template info from the API.
 */
interface TemplateInfo {
  name: string;
  lastModified: string;
  size: string;
}


/**
 * Admin component — template info and utilities.
 */
@Component({
  selector: 'app-admin',
  standalone: true,
  imports: [
    CommonModule,
    MatCardModule,
    MatIconModule,
  ],
  templateUrl: './admin.component.html',
  styleUrls: [ './admin.component.scss' ]
})
export class AdminComponent implements OnInit {

  templateInfo: TemplateInfo | null = null;


  constructor(
    private http: HttpClient
  ) {}


  /**
   * Load template info on init.
   */
  ngOnInit(): void {
    this.http.get<TemplateInfo>( '/api/template-info' ).subscribe({
      next: ( data ) => { this.templateInfo = data; },
      error: () => {},
    });
  }
}
