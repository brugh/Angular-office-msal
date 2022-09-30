/**
 * This file contains authentication parameters. Contents of this file
 * is roughly the same across other MSAL.js libraries. These parameters
 * are used to initialize Angular and MSAL Angular configurations in
 * in app.module.ts file.
 */

import { Inject, Injectable } from '@angular/core';
import { CanActivate, ActivatedRouteSnapshot, RouterStateSnapshot } from '@angular/router';
import { MsalInterceptorConfiguration, MsalGuardConfiguration, MsalBroadcastService, MsalService, MSAL_GUARD_CONFIG } from '@azure/msal-angular';
import { LogLevel, Configuration, BrowserCacheLocation, InteractionType, IPublicClientApplication, PublicClientApplication, InteractionStatus, AuthenticationResult, PopupRequest } from '@azure/msal-browser';
import { Observable, filter, switchMap, of } from 'rxjs';
import { environment } from 'src/environments/environment';

declare const Office: any;
const isIE = window.navigator.userAgent.indexOf("MSIE ") > -1 || window.navigator.userAgent.indexOf("Trident/") > -1;
const isIframe = window !== window.parent && !window.opener;

@Injectable({ providedIn: 'root' })
export class CanActivateGuard implements CanActivate {
  constructor(private msalService: MsalService, private auth: AuthService,
    private msalBroadcastService: MsalBroadcastService) { }

  canActivate(route: ActivatedRouteSnapshot, state: RouterStateSnapshot): Observable<boolean> | Promise<boolean> | boolean {
    return this.msalBroadcastService.inProgress$
      .pipe(
        filter((status: InteractionStatus) => status === InteractionStatus.None),
        switchMap(() => {
          if (this.msalService.instance.getAllAccounts().length > 0) {
            return of(true);
          }
          // this.router.navigate(['/login']);
          this.auth.login();
          return of(false);
        })
      );
  }
}


@Injectable({ providedIn: 'root' })
export class AuthService {
  constructor(
    @Inject(MSAL_GUARD_CONFIG) private msalGuardConfig: MsalGuardConfiguration,
    private authService: MsalService) { }

  login() {
    console.log("Logging in")
    if (!Office.context.ui || isIframe || this.msalGuardConfig.interactionType === InteractionType.Popup) {
      if (this.msalGuardConfig.authRequest) {
        this.authService.loginPopup({ ...this.msalGuardConfig.authRequest } as PopupRequest)
          .subscribe((response: AuthenticationResult) => {
            this.authService.instance.setActiveAccount(response.account);
          });
      } else {
        this.authService.loginPopup()
          .subscribe((response: AuthenticationResult) => {
            this.authService.instance.setActiveAccount(response.account);
          });
      }
    } else {
      Office.context.ui.displayDialogAsync(window.location.origin + '/#/login',
        { height: 50, width: 24 }, (result: any) => {
          const dialog = result.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: { message: string, type: string }) => {
            const messageFromDialog = JSON.parse(arg.message);
            if (messageFromDialog.status === 'success') {
              console.log(messageFromDialog.result);
            } else {
              console.log("Message error\n" + JSON.stringify(messageFromDialog));
            }
            console.log('Logged in');
            window.location.reload(); // For Office Add-in browser
            dialog.close();
          });
        }
      );
      // if (this.msalGuardConfig.authRequest) {
      //   this.authService.loginRedirect({ ...this.msalGuardConfig.authRequest } as RedirectRequest);
      // } else {
      //   this.authService.loginRedirect();
      // }
    }
  }

  logout() {
    // this.authService.logout();
    if (isIframe) return this.authService.logoutPopup();
    return this.authService.logoutRedirect();
  }
}


export function MSALInstanceFactory(): IPublicClientApplication {
  return new PublicClientApplication(msalConfig);
}

export function MSALInterceptorConfigFactory(): MsalInterceptorConfiguration {
  const protectedResourceMap = new Map<string, Array<string>>();

  protectedResourceMap.set(protectedResources.graphMe.endpoint, protectedResources.graphMe.scopes);
  protectedResourceMap.set(protectedResources.armTenants.endpoint, protectedResources.armTenants.scopes);

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap
  };
}

export function MSALGuardConfigFactory(): MsalGuardConfiguration {
  return {
    interactionType: InteractionType.Redirect,
    authRequest: loginRequest
  };
}

/**
 * Configuration object to be passed to MSAL instance on creation.
 * For a full list of MSAL.js configuration parameters, visit:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/configuration.md
 */
export const msalConfig: Configuration = {
  auth: {
    clientId: environment.clientId, // This is the ONLY mandatory field that you need to supply.
    authority: 'https://login.microsoftonline.com/common', // Defaults to "https://login.microsoftonline.com/common"
    redirectUri: 'https://localhost:4200/', // Points to window.location.origin. You must register this URI on Azure portal/App Registration.
  },
  cache: {
    cacheLocation: BrowserCacheLocation.LocalStorage, // Configures cache location. "sessionStorage" is more secure, but "localStorage" gives you SSO between tabs.
    storeAuthStateInCookie: isIE, // Set this to "true" if you are having issues on IE11 or Edge
  },
  system: {
    loggerOptions: {
      loggerCallback(logLevel: LogLevel, message: string) {
        console.log(message);
      },
      logLevel: LogLevel.Warning,
      piiLoggingEnabled: true
    }
  }
}

/**
 * Add here the endpoints and scopes when obtaining an access token for protected web APIs. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
export const protectedResources = {
  graphMe: {
    endpoint: "https://graph.microsoft.com/v1.0/me",
    scopes: environment.scopes,
  },
  armTenants: {
    endpoint: "https://management.azure.com/tenants",
    scopes: ["https://management.azure.com/user_impersonation"],
  }
}

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit:
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
export const loginRequest = {
  scopes: environment.scopes
};
