import { Component, Inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MAT_DIALOG_DATA, MatDialogModule, MatDialogRef } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { DraggableDialogDirective } from './draggable-dialog.directive';


/**
 * Data passed into the simple confirmation dialog.
 */
export interface SimpleConfirmDialogData {
  title: string;
  message: string;
  detail?: string;
  icon?: string;
  confirmLabel?: string;
  confirmColor?: 'primary' | 'accent' | 'warn';
}


/**
 * A lightweight confirmation dialog for yes/no decisions.
 * Returns true when confirmed, undefined/false when cancelled.
 */
@Component({
  selector: 'app-simple-confirm-dialog',
  standalone: true,
  imports: [
    CommonModule,
    MatDialogModule,
    MatButtonModule,
    MatIconModule,
    DraggableDialogDirective,
  ],
  template: `
    <h2 mat-dialog-title appDraggableDialog>
      <mat-icon *ngIf="data.icon">{{ data.icon }}</mat-icon>
      {{ data.title }}
    </h2>

    <mat-dialog-content>
      <p class="confirm-message">{{ data.message }}</p>
      <p class="confirm-detail" *ngIf="data.detail">{{ data.detail }}</p>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button mat-dialog-close>Cancel</button>
      <button mat-raised-button
              [color]="data.confirmColor || 'primary'"
              (click)="dialogRef.close( true )">
        {{ data.confirmLabel || 'Confirm' }}
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

    .confirm-message {
      color: #dee2e6;
      font-size: 0.95rem;
      margin: 0 0 8px 0;
    }

    .confirm-detail {
      color: #adb5bd;
      font-size: 0.85rem;
      margin: 0;
      line-height: 1.5;
    }

    mat-dialog-content {
      min-width: 360px;
      max-width: 480px;
    }
  `]
})
export class SimpleConfirmDialogComponent {

  constructor(
    public dialogRef: MatDialogRef<SimpleConfirmDialogComponent>,
    @Inject( MAT_DIALOG_DATA ) public data: SimpleConfirmDialogData
  ) {}
}
