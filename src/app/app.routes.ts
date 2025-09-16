import { Routes } from '@angular/router';
import { SpreadsheetWrapperComponent } from './components/spreadsheet/spreadsheet';
import { Login } from './components/login/login';
import { AuthGuard } from './guards/authGuard';

export const routes: Routes = [
  { path: 'login', component: Login },
  { path: 'spreadsheet', component: SpreadsheetWrapperComponent, canActivate: [AuthGuard] },
  { path: '**', redirectTo: 'login' }
];
