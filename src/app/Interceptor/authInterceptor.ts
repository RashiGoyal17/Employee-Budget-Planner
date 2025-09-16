import { Injectable } from '@angular/core';
import { HttpInterceptor, HttpRequest, HttpHandler } from '@angular/common/http';
import {  AuthServices } from '../services/auth-services';

@Injectable()
export class AuthInterceptor implements HttpInterceptor {
  constructor(private authService: AuthServices) {}

  intercept(req: HttpRequest<any>, next: HttpHandler) {
    const token = this.authService.getToken();
    console.log('Interceptor: Adding token to', req.url, 'Token exists:', !!token);
    if (token) {
      req = req.clone({
        setHeaders: { Authorization: `Bearer ${token}` }
      });
    }
    return next.handle(req);
  }
}