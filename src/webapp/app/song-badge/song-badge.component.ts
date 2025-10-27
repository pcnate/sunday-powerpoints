import { Component, Input } from '@angular/core';

@Component({
  selector: 'app-song-badge',
  templateUrl: './song-badge.component.html',
  styleUrls: ['./song-badge.component.scss']
})
export class SongBadgeComponent {
  @Input() song: any;
}
