import { HttpClientModule, HTTP_INTERCEPTORS } from '@angular/common/http';
import { CUSTOM_ELEMENTS_SCHEMA, NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { SharepointComponent } from './pages/sharepoint/sharepoint.component';
import { HomeComponent } from './pages/home/home.component';
import { MSALGuardConfigFactory, MSALInstanceFactory, MSALInterceptorConfigFactory } from './services/auth.service';
import { MsalBroadcastService, MsalGuard, MsalInterceptor, MsalModule, MsalRedirectComponent, MsalService,
  MSAL_GUARD_CONFIG, MSAL_INSTANCE, MSAL_INTERCEPTOR_CONFIG } from '@azure/msal-angular';
import { ProfileComponent } from './pages/profile/profile.component';
import { NavbarComponent } from './components/navbar/navbar.component';

import { allComponents, provideFluentDesignSystem } from '@fluentui/web-components';
import { ButtonComponent } from './atoms/button/button.component';
import { StoreModule } from '@ngrx/store';
import { storeReducer } from './store/store.reducers';
import { LoginComponent } from './pages/login/login.component';
import { StoreEffects } from './store/store.effects';
import { EffectsModule } from '@ngrx/effects';

provideFluentDesignSystem().register(allComponents);

@NgModule({
  declarations: [
    AppComponent,
    SharepointComponent,
    HomeComponent,
    ProfileComponent,
    NavbarComponent,
    ButtonComponent,
    LoginComponent
  ],
  imports: [
    BrowserModule,
    HttpClientModule,
    AppRoutingModule,
    MsalModule,
    StoreModule.forRoot({ store: storeReducer }, {}),
    EffectsModule.forRoot([StoreEffects])
  ],
  schemas: [ CUSTOM_ELEMENTS_SCHEMA ],
  providers: [
    {
      provide: HTTP_INTERCEPTORS,
      useClass: MsalInterceptor,
      multi: true
    },
    {
      provide: MSAL_INSTANCE,
      useFactory: MSALInstanceFactory
    },
    {
      provide: MSAL_GUARD_CONFIG,
      useFactory: MSALGuardConfigFactory
    },
    {
      provide: MSAL_INTERCEPTOR_CONFIG,
      useFactory: MSALInterceptorConfigFactory
    },
    MsalService,
    MsalGuard,
    MsalBroadcastService
  ],
  bootstrap: [AppComponent, MsalRedirectComponent]
})
export class AppModule { }
