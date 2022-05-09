import { Component, OnInit } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { Store } from '@ngrx/store';
import { StoreState } from 'src/app/store/store.state';

declare const Office: any;

@Component({
  selector: 'app-login',
  templateUrl: './login.component.html',
  styleUrls: ['./login.component.scss']
})
export class LoginComponent implements OnInit {
  loggedIn$ = this.store.select(s => s.store.loggedIn);

  constructor(private store: Store<{ store: StoreState }>, private msal: MsalService) { }

  ngOnInit(): void {
    this.home();
  }
  login() {
    this.msal.loginRedirect();
  }
  home() {
    Office.context.ui.messageParent(JSON.stringify({ status: 'success' }))
  }

}
