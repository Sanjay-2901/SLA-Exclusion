import { Component } from '@angular/core';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss'],
})
export class AppComponent {
  shouldDisable: boolean = false;

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
