import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { MsalGuard, MsalRedirectComponent } from '@azure/msal-angular';
import { CanActivateGuard } from './services/auth.service';
import { HomeComponent } from './pages/home/home.component';
import { LoginComponent } from './pages/login/login.component';
import { ProfileComponent } from './pages/profile/profile.component';
import { BrowserUtils } from '@azure/msal-browser';

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
    initialNavigation: !BrowserUtils.isInIframe() && !BrowserUtils.isInPopup() ? 'enabled' : 'disabled' // Remove this line to use Angular Universal
  })],
  exports: [RouterModule]
})
export class AppRoutingModule { }
