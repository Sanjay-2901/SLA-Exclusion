import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { NgSelectModule } from '@ng-select/ng-select';
import { ShqComponentComponent } from './shq-component/shq-component.component';
import { BlockComponentComponent } from './block-component/block-component.component';

@NgModule({
  declarations: [AppComponent, BlockComponentComponent, ShqComponentComponent],
  imports: [BrowserModule, AppRoutingModule, NgSelectModule],
  providers: [],
  bootstrap: [AppComponent],
})
export class AppModule {}
