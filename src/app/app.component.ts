import { Component } from '@angular/core';
import { SharedService } from './shared/shared.service';
import { BlockService } from './block-component/block.service';
import { GpService } from './gp/gp.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  shouldDisable: boolean = false;
  constructor(public blockService: BlockService, public gpService: GpService) {}

  isBlockLoading(event: boolean): void {
    this.shouldDisable = event;
  }

  isShqLoading(event: boolean): void {
    this.shouldDisable = event;
  }

  isGpLoading(event: boolean): void {
    this.shouldDisable = event;
  }
}
