import { Component } from '@angular/core';
import { Router, RouterModule } from '@angular/router';  // Assume you add routing
import { AuthServices } from '../../services/auth-services';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { LoginRequest } from '../../models/authModel';
import { CommonModule } from '@angular/common';


@Component({
  selector: 'app-login',
  standalone:true,
  imports: [FormsModule,CommonModule,RouterModule],
  templateUrl: './login.html',
  styleUrl: './login.scss'
})
export class Login {

  credentials: LoginRequest = { username: '', password: '' };
  error = '';

  constructor(private authService: AuthServices, private router: Router) {}

  onSubmit() {
    this.authService.login(this.credentials).subscribe({
      next: () => this.router.navigate(['/spreadsheet']),  // Route to your main component
      error: (err) => this.error = 'Login failed'
    });
  }

}
