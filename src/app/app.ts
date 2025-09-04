import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { SpreadsheetWrapperComponent } from "./components/spreadsheet/spreadsheet";

@Component({
  selector: 'app-root',
  standalone:true,
  imports: [SpreadsheetWrapperComponent],
  templateUrl: './app.html',
  styleUrl: './app.scss'
})
export class App {
  protected title = 'budget-planner';
}
