import { Component, Inject } from '@angular/core';
import { CommonModule } from '@angular/common';
import { MatDialogModule, MatDialogRef, MatDialog, MAT_DIALOG_DATA } from '@angular/material/dialog';
import { MatButtonModule } from '@angular/material/button';
import { MatIconModule } from '@angular/material/icon';
import { DraggableDialogDirective } from '../shared/draggable-dialog.directive';
import { SongSelectDialogComponent } from './song-select-dialog.component';


/**
 * Data passed into the chorus edit dialog.
 */
export interface ChorusEditDialogData {
  chorus: { name: string; number?: string };
  chorusIndex: number;
}


/**
 * Result returned when an action is taken.
 */
export interface ChorusEditDialogResult {
  action: 'change' | 'remove';
  song?: { name: string; number?: string; book?: string };
}


/**
 * Chorus edit dialog — change to a different chorus or remove.
 */
@Component({
  selector: 'app-chorus-edit-dialog',
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
      <mat-icon>music_note</mat-icon>
      Edit Chorus {{ data.chorusIndex + 1 }}
    </h2>

    <mat-dialog-content>
      <div class="current-chorus">
        <span class="current-label">Current:</span>
        <span class="current-name">{{ chorusDisplay }}</span>
      </div>
    </mat-dialog-content>

    <mat-dialog-actions align="end">
      <button mat-button mat-dialog-close>Cancel</button>
      <button mat-raised-button color="warn" (click)="remove()">
        <mat-icon>delete</mat-icon>
        Remove
      </button>
      <button mat-raised-button color="primary" (click)="change()">
        <mat-icon>swap_horiz</mat-icon>
        Change
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
      min-width: 300px;
    }

    .current-chorus {
      display: flex;
      align-items: center;
      gap: 8px;
      padding: 8px 0;
    }

    .current-label {
      color: #adb5bd;
      font-size: 0.85rem;
    }

    .current-name {
      color: #6ea8fe;
      font-size: 1rem;
      font-weight: 500;
    }
  `]
})
export class ChorusEditDialogComponent {

  constructor(
    private dialogRef: MatDialogRef<ChorusEditDialogComponent>,
    private dialog: MatDialog,
    @Inject( MAT_DIALOG_DATA ) public data: ChorusEditDialogData
  ) {}


  /**
   * Formatted display of the current chorus.
   */
  get chorusDisplay(): string {
    const c = this.data.chorus;
    if ( c.number ) return `${ c.number } - ${ c.name }`;
    return c.name;
  }


  /**
   * Remove this chorus and close.
   */
  remove(): void {
    this.dialogRef.close({ action: 'remove' } as ChorusEditDialogResult);
  }


  /**
   * Open the song select dialog to pick a replacement chorus.
   */
  change(): void {
    const selectRef = this.dialog.open( SongSelectDialogComponent, {
      width: '900px',
      maxHeight: '80vh',
      data: { currentSong: this.data.chorus, slot: 'chorus' },
    });

    selectRef.afterClosed().subscribe( ( song: { name: string; number?: string; book?: string } | undefined ) => {
      if ( !song ) return;
      this.dialogRef.close({ action: 'change', song } as ChorusEditDialogResult);
    });
  }
}
