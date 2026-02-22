import { Directive, ElementRef, OnInit, OnDestroy, NgZone } from '@angular/core';


/**
 * Directive to make a Material dialog draggable by its title bar.
 * Apply to the `<h2 mat-dialog-title>` element inside a dialog.
 *
 * Moves the `.cdk-overlay-pane` ancestor directly via CSS transform,
 * which is more reliable than CDK Drag for Material dialog overlays.
 *
 * @example
 * ```html
 * <h2 mat-dialog-title appDraggableDialog>Dialog Title</h2>
 * ```
 */
@Directive({
  selector: '[appDraggableDialog]',
  standalone: true,
})
export class DraggableDialogDirective implements OnInit, OnDestroy {

  private overlayPane: HTMLElement | null = null;
  private offsetX = 0;
  private offsetY = 0;
  private translateX = 0;
  private translateY = 0;
  private dragging = false;

  private boundMouseDown = this.onMouseDown.bind( this );
  private boundMouseMove = this.onMouseMove.bind( this );
  private boundMouseUp = this.onMouseUp.bind( this );


  constructor(
    private el: ElementRef<HTMLElement>,
    private zone: NgZone
  ) {}


  /**
   * Set up cursor styling and mouse event listener on init.
   */
  ngOnInit(): void {
    this.el.nativeElement.style.cursor = 'move';

    // Find the overlay pane ancestor
    let parent = this.el.nativeElement.parentElement;
    while ( parent ) {
      if ( parent.classList.contains( 'cdk-overlay-pane' ) ) {
        this.overlayPane = parent;
        break;
      }
      parent = parent.parentElement;
    }

    // Run outside Angular zone for performance (no change detection on mousemove)
    this.zone.runOutsideAngular( () => {
      this.el.nativeElement.addEventListener( 'mousedown', this.boundMouseDown );
    });
  }


  /**
   * Clean up all event listeners on destroy.
   */
  ngOnDestroy(): void {
    this.el.nativeElement.removeEventListener( 'mousedown', this.boundMouseDown );
    document.removeEventListener( 'mousemove', this.boundMouseMove );
    document.removeEventListener( 'mouseup', this.boundMouseUp );
  }


  /**
   * Start dragging — record the offset from mouse to current translate position.
   *
   * @param e - mousedown event
   */
  private onMouseDown( e: MouseEvent ): void {
    if ( !this.overlayPane ) return;

    // Only drag on left mouse button
    if ( e.button !== 0 ) return;

    this.dragging = true;
    this.offsetX = e.clientX - this.translateX;
    this.offsetY = e.clientY - this.translateY;

    document.addEventListener( 'mousemove', this.boundMouseMove );
    document.addEventListener( 'mouseup', this.boundMouseUp );

    e.preventDefault();
  }


  /**
   * Move the overlay pane via CSS transform while dragging.
   *
   * @param e - mousemove event
   */
  private onMouseMove( e: MouseEvent ): void {
    if ( !this.dragging || !this.overlayPane ) return;

    this.translateX = e.clientX - this.offsetX;
    this.translateY = e.clientY - this.offsetY;

    this.overlayPane.style.transform = `translate(${ this.translateX }px, ${ this.translateY }px)`;
  }


  /**
   * Stop dragging on mouse up.
   */
  private onMouseUp(): void {
    this.dragging = false;
    document.removeEventListener( 'mousemove', this.boundMouseMove );
    document.removeEventListener( 'mouseup', this.boundMouseUp );
  }
}
