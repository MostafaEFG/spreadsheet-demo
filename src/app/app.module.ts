import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';
import { BrowserAnimationsModule } from "@angular/platform-browser/animations";

import { SpreadsheetAllModule } from '@syncfusion/ej2-angular-spreadsheet';

import { IgxExcelModule } from 'igniteui-angular-excel';
import { IgxSpreadsheetModule } from 'igniteui-angular-spreadsheet';

import { SpreadSheetsModule } from "@grapecity/spread-sheets-angular";


import { AppComponent } from './app.component';
import { FormsModule } from '@angular/forms';
import { WijmoExcelComponent } from './wijmo/wijmo-excel/wijmo-excel.component';
import { WjGridModule } from '@grapecity/wijmo.angular2.grid';
import { WjGaugeModule } from '@grapecity/wijmo.angular2.gauge';
import { DataService } from './wijmo/app.data';

// ECMA 6 - using the system import method
// import { GridCore } from 'ag-grid-community';

@NgModule({
  declarations: [
    AppComponent,
    WijmoExcelComponent
  ],
  imports: [
    BrowserModule,
    FormsModule,
    SpreadsheetAllModule,
    IgxExcelModule,
    IgxSpreadsheetModule,
    BrowserAnimationsModule,
    SpreadSheetsModule,
    WjGridModule,
    WjGaugeModule,
    // Grid
  ],
  providers: [DataService],
  bootstrap: [AppComponent]
})
export class AppModule { }
