import { Component, OnInit } from '@angular/core';
import { Store } from '@ngrx/store';
import { getClaims, getPhoto } from '../store/store.actions';
import { StoreState } from '../store/store.state';

type ProfileType = {
  givenName?: string,
  surname?: string,
  userPrincipalName?: string,
  id?: string
}

@Component({
  selector: 'app-profile',
  templateUrl: './profile.component.html',
  styleUrls: ['./profile.component.scss']
})
export class ProfileComponent implements OnInit {
  profile!: ProfileType;
  displayedColumns: string[] = ['claim', 'value'];
  datasource$ = this.store.select(s => s.store.claims);
  photo$ = this.store.select(s => s.store.photo);

  constructor(private store: Store<{store: StoreState}>) { }

  ngOnInit(): void {
    this.store.dispatch(getClaims());
    this.store.dispatch(getPhoto());
    console.log('Profile')
  }


}