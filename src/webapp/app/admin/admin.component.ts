import { Component, OnInit } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { HttpClient } from '@angular/common/http';
import { MatDialogModule } from '@angular/material/dialog';
import { MatIconModule } from '@angular/material/icon';
import { MatFormFieldModule } from '@angular/material/form-field';
import { MatInputModule } from '@angular/material/input';
import { MatButtonModule } from '@angular/material/button';
import { DraggableDialogDirective } from '../shared/draggable-dialog.directive';


/**
 * Settings dialog — editable Claude processing prompt and other config.
 */
@Component({
  selector: 'app-admin',
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    MatDialogModule,
    MatIconModule,
    MatFormFieldModule,
    MatInputModule,
    MatButtonModule,
    DraggableDialogDirective,
  ],
  template: `
    <h2 mat-dialog-title appDraggableDialog>
      <mat-icon>settings</mat-icon>
      Job Settings
    </h2>

    <mat-dialog-content>
      <h3>Claude Processing Prompt</h3>
      <p class="prompt-description">
        Sent to Claude Code with the VTT transcription when processing sermon recordings.
        Saved with each job at creation time.
      </p>
      <mat-form-field appearance="outline" class="prompt-field">
        <mat-label>Prompt</mat-label>
        <textarea matInput
                  [(ngModel)]="claudePrompt"
                  rows="10"
                  placeholder="Enter the prompt template for sermon processing..."></textarea>
      </mat-form-field>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <mat-icon *ngIf="saved" class="save-check">check_circle</mat-icon>
      <button mat-button mat-dialog-close>Close</button>
      <button mat-flat-button color="primary" (click)="savePrompt()" [disabled]="saving">
        <mat-icon>save</mat-icon>
        {{ saving ? 'Saving...' : 'Save' }}
      </button>
    </mat-dialog-actions>
  `,
  styles: [`
    h2 {
      display: flex;
      align-items: center;
      gap: 8px;
      cursor: move;
    }

    mat-dialog-content {
      min-width: 500px;
    }

    h3 {
      margin: 0 0 8px;
      font-weight: 500;
      color: #dee2e6;
    }

    .prompt-description {
      font-size: 0.85rem;
      color: #6c757d;
      margin-bottom: 16px;
    }

    .prompt-field {
      width: 100%;
    }

    ::ng-deep .prompt-field .mat-mdc-form-field-subscript-wrapper {
      display: none;
    }

    .save-check {
      color: #75b798;
    }
  `]
})
export class AdminComponent implements OnInit {

  claudePrompt = '';
  saving = false;
  saved = false;
  private savedTimer: any;


  constructor(
    private http: HttpClient
  ) {}


  /**
   * Load the claude prompt on init.
   */
  ngOnInit(): void {
    this.http.get<{ value: string }>( '/api/settings/claude-prompt' ).subscribe({
      next: ( data ) => { this.claudePrompt = data.value; },
      error: () => {},
    });
  }


  /**
   * Save the claude prompt to the server.
   */
  savePrompt(): void {
    this.saving = true;
    this.saved = false;
    clearTimeout( this.savedTimer );

    this.http.put( '/api/settings/claude-prompt', { value: this.claudePrompt } ).subscribe({
      next: () => {
        this.saving = false;
        this.saved = true;
        this.savedTimer = setTimeout( () => { this.saved = false; }, 3000 );
      },
      error: () => { this.saving = false; },
    });
  }
}
