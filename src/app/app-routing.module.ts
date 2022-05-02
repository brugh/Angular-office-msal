import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { MsalGuard, MsalRedirectComponent } from '@azure/msal-angular';
import { CanActivateGuard } from './services/auth.service';
import { HomeComponent } from './home/home.component';
import { LoginComponent } from './login/login.component';
import { ProfileComponent } from './profile/profile.component';

const routes: Routes = [
  { path: 'profile', component: ProfileComponent, canActivate: [CanActivateGuard, MsalGuard] },
  { path: 'login', component: LoginComponent, canActivate: [MsalGuard]},
  { // Needed for hash routing
    path: 'error', component: HomeComponent
  },
  { // Needed for hash routing
    path: 'state', component: HomeComponent
  },
  { // Needed for hash routing
    path: 'code', component: HomeComponent
  },
  { path: 'auth', component: MsalRedirectComponent },
  { path: '', component: HomeComponent },
  { path: '**', redirectTo: '', pathMatch: 'full' }
];

const isIframe = window !== window.parent && !window.opener;

@NgModule({
  imports: [RouterModule.forRoot(routes, { 
    useHash: true, 
    enableTracing: false,
    initialNavigation: 'enabled' // !isIframe ? 'enabled' : 'disabled'
  })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
