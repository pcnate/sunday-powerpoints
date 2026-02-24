import { Component, Inject, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MatDialogModule, MatDialogRef, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatRadioModule } from '@angular/material/radio';
import { MatAutocompleteModule } from '@angular/material/autocomplete';
import { MatProgressSpinnerModule } from '@angular/material/progress-spinner';
import { DraggableDialogDirective } from '../shared/draggable-dialog.directive';


/**
 * Data passed into the sermon picker dialog.
 */
export interface SermonPickerDialogData {
  folderDate: string;
  currentPptxFile: string | null;
  currentTitle: string | null;
  currentSpeaker: string | null;
}


/**
 * Result returned when saving and closing.
 */
export interface SermonPickerDialogResult {
  pptxFile: string | null;
  title: string | null;
  speaker: string | null;
}


/**
 * Sermon picker dialog — select a sermon PPTX file from the folder,
 * set a title, and assign speaker(s).
 */
@Component({
  selector: 'app-sermon-picker-dialog',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatDialogModule,
    MatButtonModule,
    MatIconModule,
    MatFormFieldModule,
    MatInputModule,
    MatRadioModule,
    MatAutocompleteModule,
    MatProgressSpinnerModule,
    DraggableDialogDirective,
  ],
  template: `
    <h2 mat-dialog-title appDraggableDialog>
      <mat-icon>church</mat-icon>
      Sermon — {{ data.folderDate }}
    </h2>

    <mat-dialog-content>
      <div *ngIf="loading" class="loading">
        <mat-spinner diameter="30"></mat-spinner>
      </div>

      <ng-container *ngIf="!loading">
        <div class="file-section" *ngIf="pptxFiles.length > 0; else noFiles">
          <label class="section-label">Presentation File</label>
          <mat-radio-group [(ngModel)]="selectedFile" class="file-radio-group">
            <mat-radio-button [value]="''" class="file-option">
              <span class="file-none">None</span>
            </mat-radio-button>
            <mat-radio-button *ngFor="let file of pptxFiles" [value]="file" class="file-option">
              <mat-icon class="file-icon">slideshow</mat-icon>
              {{ file }}
            </mat-radio-button>
          </mat-radio-group>
        </div>
        <ng-template #noFiles>
          <p class="no-files-message">No sermon presentations found in this folder.</p>
        </ng-template>

        <mat-form-field appearance="outline" class="full-width">
          <mat-label>Sermon Title</mat-label>
          <input matInput [(ngModel)]="title" placeholder="e.g. The Good Shepherd">
        </mat-form-field>

        <mat-form-field appearance="outline" class="full-width">
          <mat-label>Speaker(s)</mat-label>
          <input matInput
                 [(ngModel)]="speaker"
                 [matAutocomplete]="speakerAuto"
                 placeholder="e.g. Pastor John">
          <mat-autocomplete #speakerAuto="matAutocomplete">
            <mat-option *ngFor="let s of filteredSpeakers" [value]="s">
              {{ s }}
            </mat-option>
          </mat-autocomplete>
        </mat-form-field>
      </ng-container>

      <div *ngIf="statusMessage" class="status-message" [class.error]="hasError">
        {{ statusMessage }}
      </div>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button (click)="dialogRef.close()">Cancel</button>
      <button mat-raised-button color="primary" (click)="save()" [disabled]="saving || loading">
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

    .section-label {
      display: block;
      font-size: 0.8rem;
      color: #adb5bd;
      margin-bottom: 8px;
    }

    .file-section {
      margin-bottom: 16px;
    }

    .file-radio-group {
      display: flex;
      flex-direction: column;
      gap: 4px;
      margin-bottom: 12px;
    }

    .file-option {
      font-size: 0.9rem;
    }

    .file-icon {
      font-size: 18px;
      width: 18px;
      height: 18px;
      vertical-align: middle;
      margin-right: 4px;
      color: #6ea8fe;
    }

    .file-none {
      color: #6c757d;
      font-style: italic;
    }

    .no-files-message {
      color: #6c757d;
      font-size: 0.9rem;
      margin: 0 0 16px;
    }

    .full-width {
      width: 100%;
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
      min-width: 400px;
      overflow: visible;
    }
  `]
})
export class SermonPickerDialogComponent implements OnInit {

  pptxFiles: string[] = [];
  selectedFile = '';
  title = '';
  speaker = '';
  loading = true;
  saving = false;
  statusMessage = '';
  hasError = false;

  /** All saved speakers from the DB. */
  allSpeakers: string[] = [];


  constructor(
    public dialogRef: MatDialogRef<SermonPickerDialogComponent>,
    @Inject( MAT_DIALOG_DATA ) public data: SermonPickerDialogData,
    private http: HttpClient
  ) {
    this.selectedFile = data.currentPptxFile || '';
    this.title = data.currentTitle || '';
    this.speaker = data.currentSpeaker || '';
  }


  /**
   * Filtered speaker suggestions based on current input.
   */
  get filteredSpeakers(): string[] {
    if ( !this.speaker.trim() ) return this.allSpeakers;
    const lower = this.speaker.toLowerCase();
    return this.allSpeakers.filter( s => s.toLowerCase().includes( lower ) );
  }


  /**
   * Load available PPTX files and speakers on dialog open.
   */
  ngOnInit(): void {
    this.http.get<{ files: string[] }>( `/api/sermon-pptx-files?folder=${ this.data.folderDate }` ).subscribe({
      next: ( res ) => {
        this.pptxFiles = res.files || [];
        this.loading = false;
      },
      error: () => {
        this.pptxFiles = [];
        this.loading = false;
      },
    });

    this.http.get<{ id: number; name: string }[]>( '/api/speakers' ).subscribe({
      next: ( speakers ) => {
        this.allSpeakers = speakers.map( s => s.name );
      },
      error: () => {
        this.allSpeakers = [];
      },
    });
  }


  /**
   * Save sermon info to the server, then close with the result.
   * If the speaker name is new, it is automatically added to the speakers list.
   */
  save(): void {
    this.saving = true;
    this.statusMessage = '';
    this.hasError = false;

    const speakerName = this.speaker.trim() || null;

    // If speaker is new and non-empty, add to the saved list
    if ( speakerName && !this.allSpeakers.some( s => s.toLowerCase() === speakerName.toLowerCase() ) ) {
      this.http.post( '/api/speakers', { name: speakerName } ).subscribe();
    }

    this.http.put( '/api/sermon-info', {
      folder: this.data.folderDate,
      pptxFile: this.selectedFile || null,
      title: this.title.trim() || null,
      speaker: speakerName,
    }).subscribe({
      next: () => {
        this.saving = false;
        this.dialogRef.close({
          pptxFile: this.selectedFile || null,
          title: this.title.trim() || null,
          speaker: speakerName,
        } as SermonPickerDialogResult);
      },
      error: ( err ) => {
        this.statusMessage = `Save failed: ${ err.message || 'Unknown error' }`;
        this.hasError = true;
        this.saving = false;
      },
    });
  }
}
