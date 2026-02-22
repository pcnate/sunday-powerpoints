import { Component, Inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MatDialogModule, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { DraggableDialogDirective } from '../shared/draggable-dialog.directive';


/**
 * Data passed into the verse dialog.
 */
export interface VerseDialogData {
  year: number;
  month: number;
  monthName: string;
  reference: string;
  text: string;
}


/**
 * Result returned when saving and closing.
 */
export interface VerseDialogResult {
  reference: string;
  text: string;
}


/**
 * Bible verse editor dialog — edit reference and full text for a month.
 */
@Component({
  selector: 'app-verse-dialog',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatDialogModule,
    MatButtonModule,
    MatIconModule,
    MatFormFieldModule,
    MatInputModule,
    DraggableDialogDirective,
  ],
  template: `
    <h2 mat-dialog-title appDraggableDialog>
      <mat-icon>menu_book</mat-icon>
      Bible Verse — {{ data.monthName }} {{ data.year }}
    </h2>

    <mat-dialog-content>
      <mat-form-field appearance="outline" class="full-width">
        <mat-label>Verse Reference</mat-label>
        <input matInput [(ngModel)]="reference" placeholder="e.g. Romans 8:28">
      </mat-form-field>

      <a *ngIf="bibleGatewayUrl()" [href]="bibleGatewayUrl()" target="_blank" rel="noopener"
         mat-button class="gateway-link">
        <mat-icon>open_in_new</mat-icon>
        Open in Bible Gateway (NKJV)
      </a>

      <textarea
        [(ngModel)]="verseText"
        class="verse-textarea"
        placeholder="Paste the verse text here..."
        rows="6"
      ></textarea>

      <div *ngIf="statusMessage" class="status-message" [class.error]="hasError">
        {{ statusMessage }}
      </div>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button (click)="dialogRef.close()">Cancel</button>
      <button mat-raised-button color="primary" (click)="save()" [disabled]="saving">
        <mat-icon>save</mat-icon>
        Save
      </button>
    </mat-dialog-actions>
  `,
  styles: [`
    .full-width {
      width: 100%;
    }

    .gateway-link {
      color: #6ea8fe;
      margin-bottom: 12px;
      display: inline-flex;
      align-items: center;
      gap: 4px;
    }

    .verse-textarea {
      width: 100%;
      min-height: 120px;
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

    .verse-textarea:focus {
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

    mat-dialog-content {
      min-width: 450px;
      overflow: visible;
    }
  `]
})
export class VerseDialogComponent {

  reference: string;
  verseText: string;
  saving = false;
  statusMessage = '';
  hasError = false;


  constructor(
    public dialogRef: MatDialogRef<VerseDialogComponent>,
    @Inject( MAT_DIALOG_DATA ) public data: VerseDialogData,
    private http: HttpClient
  ) {
    this.reference = data.reference || '';
    this.verseText = data.text || '';
  }


  /**
   * Build the Bible Gateway NKJV URL for the current reference.
   *
   * @returns full URL string or empty string if no reference
   */
  bibleGatewayUrl(): string {
    if ( !this.reference.trim() ) return '';
    return `https://www.biblegateway.com/passage/?search=${ encodeURIComponent( this.reference.trim() ) }&version=NKJV`;
  }


  /**
   * Save verse reference and text to the server, then close.
   */
  save(): void {
    this.saving = true;
    this.statusMessage = '';
    this.hasError = false;

    this.http.post( '/api/verse', {
      year: this.data.year,
      month: this.data.month,
      reference: this.reference,
      text: this.verseText,
    }).subscribe({
      next: () => {
        this.saving = false;
        this.dialogRef.close({
          reference: this.reference,
          text: this.verseText,
        } as VerseDialogResult);
      },
      error: ( err ) => {
        this.statusMessage = `Save failed: ${ err.message || 'Unknown error' }`;
        this.hasError = true;
        this.saving = false;
      },
    });
  }
}
