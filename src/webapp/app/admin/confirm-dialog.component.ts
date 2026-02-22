import { Component, Inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MAT_DIALOG_DATA, MatDialogModule, MatDialogRef } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';
import { MatCheckboxModule } from '@angular/material/checkbox';
import { MatIconModule } from '@angular/material/icon';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { DraggableDialogDirective } from '../shared/draggable-dialog.directive';


/**
 * Data passed into the confirmation dialog.
 */
export interface ConfirmDialogData {
  title: string;
  apiEndpoint: string;
  month: number;
  monthName: string;
  year: number;
  showOverwrite?: boolean;
}


/**
 * Result returned when the dialog is closed.
 */
export interface ConfirmDialogResult {
  executed: boolean;
}


/**
 * Admin action dialog with month/year selection, write checkbox,
 * Run button, and script output display.
 * Matches the old site's modal behavior — stays open and shows output.
 */
@Component({
  selector: 'app-confirm-dialog',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatDialogModule,
    MatButtonModule,
    MatCheckboxModule,
    MatIconModule,
    MatProgressSpinnerModule,
    DraggableDialogDirective,
  ],
  template: `
    <h2 mat-dialog-title appDraggableDialog>
      {{ data.title }}
    </h2>

    <mat-dialog-content>
      <p class="confirm-detail">
        <strong>Month:</strong> {{ data.monthName }} {{ data.year }}
      </p>

      <mat-checkbox [(ngModel)]="writeEnabled" color="warn" class="write-checkbox">
        Write (actually create folders/files)
      </mat-checkbox>

      <mat-checkbox *ngIf="data.showOverwrite" [(ngModel)]="overwriteEnabled" color="warn" class="write-checkbox">
        Overwrite existing presentations
      </mat-checkbox>

      <div class="run-row">
        <button mat-raised-button color="primary"
                [disabled]="running || runBlocked"
                (click)="run()">
          <mat-icon *ngIf="!running">play_arrow</mat-icon>
          <mat-spinner *ngIf="running" diameter="18" class="inline-spinner"></mat-spinner>
          Run
        </button>
      </div>

      <pre *ngIf="output" class="output-block"
           [class.error-output]="hasError">{{ output }}</pre>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button (click)="close()">Close</button>
    </mat-dialog-actions>
  `,
  styles: [`
    .confirm-detail {
      color: #adb5bd;
      margin-bottom: 12px;
    }

    .write-checkbox {
      display: block;
      margin-bottom: 16px;
    }

    .run-row {
      margin-bottom: 12px;
    }

    .inline-spinner {
      display: inline-block;
      margin-right: 4px;
      vertical-align: middle;
    }

    .output-block {
      background: #1a3a1a;
      color: #c8e6c9;
      border: 1px solid #2e7d32;
      border-radius: 4px;
      padding: 12px;
      font-family: 'Consolas', 'Courier New', monospace;
      font-size: 0.85rem;
      white-space: pre-wrap;
      word-wrap: break-word;
      max-height: 300px;
      overflow-y: auto;
    }

    .error-output {
      background: #3a1a1a;
      color: #ef9a9a;
      border-color: #c62828;
    }

    mat-dialog-content {
      min-width: 500px;
    }
  `]
})
export class ConfirmDialogComponent {

  writeEnabled = false;
  overwriteEnabled = false;
  running = false;
  output = '';
  hasError = false;


  constructor(
    private http: HttpClient,
    private dialogRef: MatDialogRef<ConfirmDialogComponent>,
    @Inject( MAT_DIALOG_DATA ) public data: ConfirmDialogData
  ) {}


  /**
   * True when presentations already exist and overwrite is not checked.
   */
  get runBlocked(): boolean {
    return !!this.data.showOverwrite && !this.overwriteEnabled;
  }


  /**
   * Execute the API call and display the output.
   */
  run(): void {
    this.running = true;
    this.output = '';
    this.hasError = false;

    this.http.post<{ success: boolean; message?: string; error?: string }>( this.data.apiEndpoint, {
      month: this.data.month,
      year: this.data.year,
      write: this.writeEnabled,
      overwrite: this.overwriteEnabled,
    }).subscribe({
      next: ( res ) => {
        this.running = false;
        if ( res.success ) {
          this.output = res.message || 'Done.';
        } else {
          this.hasError = true;
          this.output = res.error || 'Unknown error';
        }
      },
      error: ( err ) => {
        this.running = false;
        this.hasError = true;
        this.output = 'Request failed: ' + ( err.message || 'Unknown error' );
      },
    });
  }


  /**
   * Close the dialog.
   */
  close(): void {
    this.dialogRef.close({ executed: !!this.output } as ConfirmDialogResult);
  }
}
