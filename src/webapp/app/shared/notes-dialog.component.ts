import { Component, Inject, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MatDialogModule, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { DraggableDialogDirective } from './draggable-dialog.directive';


/**
 * Data passed to the notes dialog.
 */
export interface NotesDialogData {
  folderName: string;
}


/**
 * Notes editor dialog — loads and saves notes for a Sunday folder.
 */
@Component({
  selector: 'app-notes-dialog',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatDialogModule,
    MatButtonModule,
    MatIconModule,
    MatProgressSpinnerModule,
    DraggableDialogDirective,
  ],
  template: `
    <h2 mat-dialog-title appDraggableDialog>
      <mat-icon>description</mat-icon>
      Notes for {{ data.folderName }}
    </h2>

    <mat-dialog-content>
      <div *ngIf="loading" class="loading">
        <mat-spinner diameter="30"></mat-spinner>
      </div>

      <textarea
        *ngIf="!loading"
        [(ngModel)]="notesText"
        class="notes-textarea"
        placeholder="Enter notes for this week..."
        rows="15"
      ></textarea>

      <div *ngIf="statusMessage" class="status-message" [class.error]="hasError">
        {{ statusMessage }}
      </div>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button (click)="dialogRef.close()">Close</button>
      <button mat-raised-button color="primary" (click)="save()" [disabled]="saving">
        <mat-icon>save</mat-icon>
        Save
      </button>
    </mat-dialog-actions>
  `,
  styles: [`
    .loading {
      display: flex;
      justify-content: center;
      padding: 24px;
    }

    .notes-textarea {
      width: 100%;
      min-height: 300px;
      padding: 12px;
      font-family: monospace;
      font-size: 0.9rem;
      background: #1a1d21;
      color: #e0e0e0;
      border: 1px solid #495057;
      border-radius: 4px;
      resize: vertical;
      box-sizing: border-box;
    }

    .notes-textarea:focus {
      outline: none;
      border-color: #6ea8fe;
    }

    .status-message {
      margin-top: 8px;
      font-size: 0.85rem;
      color: #4caf50;
    }

    .status-message.error {
      color: #f44336;
    }

    h2 {
      display: flex;
      align-items: center;
      gap: 8px;
      cursor: move;
    }
  `]
})
export class NotesDialogComponent implements OnInit {

  notesText = '';
  loading = true;
  saving = false;
  statusMessage = '';
  hasError = false;


  constructor(
    public dialogRef: MatDialogRef<NotesDialogComponent>,
    @Inject( MAT_DIALOG_DATA ) public data: NotesDialogData,
    private http: HttpClient
  ) {}


  /**
   * Load existing notes on dialog open.
   */
  ngOnInit(): void {
    this.http.get( `/api/notes?folder=${ encodeURIComponent( this.data.folderName ) }`, {
      responseType: 'text'
    }).subscribe({
      next: ( text ) => {
        this.notesText = text;
        this.loading = false;
      },
      error: () => {
        this.notesText = '';
        this.loading = false;
      },
    });
  }


  /**
   * Save notes to the server.
   */
  save(): void {
    this.saving = true;
    this.statusMessage = '';
    this.hasError = false;

    this.http.post(
      `/api/notes?folder=${ encodeURIComponent( this.data.folderName ) }`,
      this.notesText,
      { headers: { 'Content-Type': 'text/plain' }, responseType: 'text' }
    ).subscribe({
      next: () => {
        this.statusMessage = 'Saved successfully';
        this.saving = false;
      },
      error: ( err ) => {
        this.statusMessage = `Save failed: ${ err.message || 'Unknown error' }`;
        this.hasError = true;
        this.saving = false;
      },
    });
  }
}
