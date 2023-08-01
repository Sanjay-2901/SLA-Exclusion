import { Component } from '@angular/core';
import { BlockService } from './block-component/block.service';
import { ShqService } from './shq-component/shq-service.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  shouldDisable: boolean = false;
  constructor(
    public blockService: BlockService,
    public shqService: ShqService
  ) {}

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
