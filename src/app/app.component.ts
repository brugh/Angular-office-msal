import { Component, OnInit, OnDestroy } from '@angular/core';
import { MsalService, MsalBroadcastService } from '@azure/msal-angular';
import { AuthenticationResult, EventMessage, EventType, InteractionStatus } from '@azure/msal-browser';
import { Store } from '@ngrx/store';
import { Subject } from 'rxjs';
import { filter, takeUntil } from 'rxjs/operators';
import { AuthService } from './services/auth.service';
import { setLoggedIn, setClaims } from './store/store.actions';
import { StoreState } from './store/store.state';

declare const Office: any;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent implements OnInit, OnDestroy {
  private readonly _destroying$ = new Subject<void>();
  private readonly _destroying2$ = new Subject<void>();
  loggedIn$ = this.store.select(s => s.store.loggedIn);

  constructor(
    private authService: MsalService, public auth: AuthService, private msalBroadcastService: MsalBroadcastService,
    private store: Store<{ store: StoreState }>
  ) { }

  ngOnInit(): void {
    /**
     * You can subscribe to MSAL events as shown below. For more info,
     * visit: https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-angular/docs/v2-docs/events.md
     */
    this.msalBroadcastService.inProgress$
      .pipe(
        filter((status: InteractionStatus) => status === InteractionStatus.None),
        takeUntil(this._destroying$)
      )
      .subscribe((ev) => {
        this.setLoginDisplay();
        this.checkAndSetActiveAccount();
        this.getClaims(this.authService.instance.getActiveAccount()?.idTokenClaims)
      });

    this.msalBroadcastService.msalSubject$
      .pipe(
        filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS),
        takeUntil(this._destroying2$)
      )
      .subscribe((ev: EventMessage) => {
        const payload = ev.payload as AuthenticationResult;
        this.authService.instance.setActiveAccount(payload?.account);
      });
  }

  getClaims(tokenClaims: any) {
    console.log(tokenClaims);
    this.store.dispatch(
      setClaims({
        claims: [
          { id: 3, claim: "Display Name", value: tokenClaims ? tokenClaims['name'] : null },
          { id: 4, claim: "User Principal Name (UPN)", value: tokenClaims ? tokenClaims['preferred_username'] : null },
          { id: 5, claim: "OID", value: tokenClaims ? tokenClaims['oid'] : null }
        ]
      })
    );
  }

  checkAndSetActiveAccount() {
    /**
     * If no active account set but there are accounts signed in, sets first account to active account
     * To use active account set here, subscribe to inProgress$ first in your component
     * Note: Basic usage demonstrated. Your app may require more complicated account selection logic
     */
    let activeAccount = this.authService.instance.getActiveAccount();

    if (!activeAccount && this.authService.instance.getAllAccounts().length > 0) {
      let accounts = this.authService.instance.getAllAccounts();
      this.authService.instance.setActiveAccount(accounts[0]);
    }
  }

  setLoginDisplay() {
    const _L = this.authService.instance.getAllAccounts().length > 0;
    this.store.dispatch(setLoggedIn({ loggedIn: _L }))
  }

  // unsubscribe to events when component is destroyed
  ngOnDestroy(): void {
    this._destroying$.next(undefined);
    this._destroying$.complete();
    this._destroying2$.next(undefined);
    this._destroying2$.complete();
  }
}
