import { NgModule }      from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { SpreadSheetsModule } from './lib/gc.spread.sheets.angular.12.0.0';
import { AppComponent }  from './app.component';

@NgModule({
  imports:      [ BrowserModule, SpreadSheetsModule ],
  declarations: [ AppComponent ],
  bootstrap:    [ AppComponent ]
})
export class AppModule { }
