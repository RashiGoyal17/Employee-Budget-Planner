import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { Observable, BehaviorSubject } from 'rxjs';
import { map } from 'rxjs/operators';
import { LoginRequest, AuthResponse } from '../models/authModel';

@Injectable({
  providedIn: 'root'
})
export class AuthServices {

  private readonly apiRoot = 'https://localhost:7225/api';
  private tokenSubject = new BehaviorSubject<string | null>(localStorage.getItem('token'));
  public token$ = this.tokenSubject.asObservable();

  constructor(private http: HttpClient) {}

  login(credentials: LoginRequest): Observable<AuthResponse> {
    return this.http.post<AuthResponse>(`${this.apiRoot}/auth/login`, credentials).pipe(
      map(res => {
        localStorage.setItem('token', res.token);
        this.tokenSubject.next(res.token);
        return res;
      })
    );
  }

  logout(): void {
    localStorage.removeItem('token');
    this.tokenSubject.next(null);
  }

  isLoggedIn(): boolean {
    return !!this.tokenSubject.value;
  }

  getToken(): string | null {
    return this.tokenSubject.value;
  }
  
}
