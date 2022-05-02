import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';

@Component({
  selector: 'app-navbar',
  templateUrl: './navbar.component.html',
  styleUrls: ['./navbar.component.scss']
})
export class NavbarComponent implements OnInit {
  @Input() loggedIn: boolean | null = false;
  @Output() login = new EventEmitter<any>();
  @Output() logout = new EventEmitter<any>();

  constructor() { }

  ngOnInit(): void {
  }

}
