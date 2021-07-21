import { CUSTOM_ELEMENTS_SCHEMA, NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { FormsModule } from '@angular/forms';
import { AppComponent } from './app.component';
import { SheetJSComponent } from './sheet/sheet.component';


@NgModule({
  imports:      [ BrowserModule, 
                  FormsModule],
  declarations: [ AppComponent, SheetJSComponent ],
  bootstrap:    [ AppComponent ],
  providers: [

  ],
  schemas: [CUSTOM_ELEMENTS_SCHEMA],
})
export class AppModule { }
