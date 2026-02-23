import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MatDialogModule, MatDialogRef } from '@angular/material/dialog';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatSelectModule } from '@angular/material/select';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { DraggableDialogDirective } from '../shared/draggable-dialog.directive';


/**
 * A book from the API.
 */
interface Book {
  id: number;
  name: string;
}


/**
 * Dialog for creating a new song file from the template.
 * Lets the user pick a book, enter an optional number, required name, and optional CCLI.
 */
@Component({
  selector: 'app-create-song-dialog',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatDialogModule,
    MatFormFieldModule,
    MatInputModule,
    MatSelectModule,
    MatButtonModule,
    MatIconModule,
    MatProgressSpinnerModule,
    DraggableDialogDirective,
  ],
  template: `
    <h2 mat-dialog-title appDraggableDialog>
      <mat-icon>add_circle</mat-icon> Create New Song
    </h2>

    <mat-dialog-content>
      <div class="form-fields">
        <mat-form-field appearance="outline" class="full-width">
          <mat-label>Book / Source</mat-label>
          <mat-select [(ngModel)]="selectedBook" (selectionChange)="onBookChange()">
            <mat-option *ngFor="let b of books" [value]="b.name">{{ b.name }}</mat-option>
            <mat-option value="__other__">Other…</mat-option>
          </mat-select>
        </mat-form-field>

        <mat-form-field appearance="outline" class="full-width" *ngIf="selectedBook === '__other__'">
          <mat-label>New Book Name</mat-label>
          <input matInput [(ngModel)]="newBookName" placeholder="e.g. Special Songs">
        </mat-form-field>

        <div class="number-name-row">
          <mat-form-field appearance="outline" class="number-field">
            <mat-label>Number</mat-label>
            <input matInput [(ngModel)]="songNumber" placeholder="e.g. 123" maxlength="5">
          </mat-form-field>

          <mat-form-field appearance="outline" class="name-field">
            <mat-label>Song Name</mat-label>
            <input matInput [(ngModel)]="songName" placeholder="e.g. Amazing Grace" required>
          </mat-form-field>
        </div>

        <mat-form-field appearance="outline" class="full-width">
          <mat-label>CCLI Number (optional)</mat-label>
          <input matInput [(ngModel)]="ccli" placeholder="e.g. 4669076">
        </mat-form-field>
      </div>

      <div class="preview" *ngIf="previewPath">
        <mat-icon>folder</mat-icon>
        <span>{{ previewPath }}</span>
      </div>

      <div class="error-message" *ngIf="errorMessage">
        <mat-icon>error</mat-icon>
        <span>{{ errorMessage }}</span>
      </div>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button mat-dialog-close>Cancel</button>
      <button mat-raised-button color="primary"
              [disabled]="!canCreate || saving"
              (click)="create()">
        <mat-spinner *ngIf="saving" diameter="18" class="inline-spinner"></mat-spinner>
        {{ saving ? 'Creating...' : 'Create' }}
      </button>
    </mat-dialog-actions>
  `,
  styles: [`
    h2 {
      display: flex;
      align-items: center;
      gap: 8px;
    }

    mat-dialog-content {
      min-width: 450px;
      max-width: 550px;
    }

    .form-fields {
      display: flex;
      flex-direction: column;
      gap: 4px;
    }

    .full-width {
      width: 100%;
    }

    .number-name-row {
      display: flex;
      gap: 12px;
    }

    .number-field {
      width: 120px;
      flex-shrink: 0;
    }

    .name-field {
      flex: 1;
    }

    .preview {
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 8px 12px;
      border-radius: 4px;
      background: rgba( 110, 168, 254, 0.08 );
      color: #6ea8fe;
      font-size: 0.85rem;
      font-family: monospace;
      margin-top: 4px;
    }

    .preview mat-icon {
      font-size: 18px;
      width: 18px;
      height: 18px;
    }

    .error-message {
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 8px 12px;
      border-radius: 4px;
      background: rgba( 244, 67, 54, 0.1 );
      color: #f44336;
      font-size: 0.85rem;
      margin-top: 8px;
    }

    .error-message mat-icon {
      font-size: 18px;
      width: 18px;
      height: 18px;
    }

    .inline-spinner {
      display: inline-block;
      margin-right: 6px;
    }
  `]
})
export class CreateSongDialogComponent implements OnInit {

  books: Book[] = [];
  selectedBook = '';
  newBookName = '';
  songNumber = '';
  songName = '';
  ccli = '';
  saving = false;
  errorMessage = '';


  constructor(
    private http: HttpClient,
    private dialogRef: MatDialogRef<CreateSongDialogComponent>
  ) {}


  /**
   * Load available books on init.
   */
  ngOnInit(): void {
    this.http.get<Book[]>( '/api/books' ).subscribe({
      next: ( data ) => { this.books = data; },
      error: () => {},
    });
  }


  /**
   * Clear new book name when switching away from "Other…".
   */
  onBookChange(): void {
    if ( this.selectedBook !== '__other__' ) {
      this.newBookName = '';
    }
    this.errorMessage = '';
  }


  /**
   * The resolved book name (selected or newly typed).
   */
  get resolvedBook(): string {
    return this.selectedBook === '__other__'
      ? this.newBookName.trim()
      : this.selectedBook;
  }


  /**
   * Whether all required fields are filled.
   */
  get canCreate(): boolean {
    return !!this.resolvedBook && !!this.songName.trim();
  }


  /**
   * Preview of the file path that will be created.
   */
  get previewPath(): string {
    if ( !this.resolvedBook || !this.songName.trim() ) return '';
    const num = this.songNumber.trim();
    const name = this.songName.trim();
    const filename = num ? `${ num } - ${ name }.pptx` : `${ name }.pptx`;
    return `${ this.resolvedBook } / ${ filename }`;
  }


  /**
   * Create the song file on the server.
   */
  create(): void {
    if ( !this.canCreate || this.saving ) return;

    this.saving = true;
    this.errorMessage = '';

    // If "Other…" was selected, create the book first
    const bookStep = this.selectedBook === '__other__'
      ? this.http.post<Book>( '/api/books', { name: this.newBookName.trim() } )
      : null;

    const doCreate = () => {
      this.http.post<{ success: boolean; song: any }>( '/api/songs/create', {
        book: this.resolvedBook,
        name: this.songName.trim(),
        number: this.songNumber.trim() || undefined,
        ccli: this.ccli.trim() || undefined,
      }).subscribe({
        next: ( result ) => {
          this.saving = false;
          this.dialogRef.close( result.song );
        },
        error: ( err ) => {
          this.saving = false;
          this.errorMessage = err.error?.error || 'Failed to create song';
        },
      });
    };

    if ( bookStep ) {
      bookStep.subscribe({
        next: () => doCreate(),
        error: ( err ) => {
          this.saving = false;
          this.errorMessage = err.error?.error || 'Failed to create book';
        },
      });
    } else {
      doCreate();
    }
  }
}
