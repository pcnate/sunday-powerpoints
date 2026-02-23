import { Routes } from '@angular/router';

export const routes: Routes = [
  { path: '', redirectTo: 'planning', pathMatch: 'full' },
  { path: 'planning', children: [] },
  { path: 'library', children: [] },
  { path: 'outstanding', children: [] },
  { path: 'jobs', children: [] },
  { path: '**', redirectTo: 'planning' },
];
