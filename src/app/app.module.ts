import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { KeysPipe } from './keys.pipe';
import { AppComponent } from './app.component';
import { ExcelService } from './excel.service';
import { AppRoutingModule } from './app-routing.module';


@NgModule({
  declarations: [
    AppComponent, KeysPipe
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
  ],
  providers: [ExcelService],
  bootstrap: [AppComponent]
})
export class AppModule { }
