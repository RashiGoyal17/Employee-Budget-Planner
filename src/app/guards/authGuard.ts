import { inject } from '@angular/core';
import { CanActivateFn, Router } from '@angular/router';
import { AuthServices } from '../services/auth-services';

export const AuthGuard: CanActivateFn = () => {
  const authService = inject(AuthServices);
  const router = inject(Router);
  if (authService.isLoggedIn()) return true;
  router.navigate(['/login']);
  return false;
};