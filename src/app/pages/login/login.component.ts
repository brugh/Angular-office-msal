import { Component, OnInit } from '@angular/core';

declare const Office: any;

@Component({
  selector: 'app-login',
  templateUrl: './login.component.html',
  styleUrls: ['./login.component.scss']
})
export class LoginComponent implements OnInit {

  constructor() {
  }

  ngOnInit(): void {
    this.home();
  }

  home() {
    Office.context.ui.messageParent(JSON.stringify({ status: 'success' }))
  }

}
