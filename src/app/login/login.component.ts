import { Component, OnInit } from '@angular/core';
import { EventMessage, EventType, AuthenticationResult } from '@azure/msal-browser';
import { filter, Subject, takeUntil } from 'rxjs';

declare const Office: any;

@Component({
  selector: 'app-login',
  templateUrl: './login.component.html',
  styleUrls: ['./login.component.scss']
})
export class LoginComponent implements OnInit {
  msalBroadcastService: any;
  authService: any;
  private readonly _destroying$ = new Subject<void>();

  constructor() {
  }

  ngOnInit(): void {
    // this.msalBroadcastService.msalSubject$
    //   .pipe(
    //     filter((msg: EventMessage) => msg.eventType === EventType.LOGIN_SUCCESS),
    //     takeUntil(this._destroying$)
    //   )
    //   .subscribe((ev: EventMessage) => {
    //     console.log('Subject', ev);
    //     const payload = ev.payload as AuthenticationResult;
    //     this.authService.instance.setActiveAccount(payload?.account);
    //   });
    this.home();
  }

  home() {
    Office.context.ui.messageParent(JSON.stringify({ status: 'success' }))
  }

  // ngOnDestroy(): void {
  //   this._destroying$.next(undefined);
  //   this._destroying$.complete(); 
  // }
}
