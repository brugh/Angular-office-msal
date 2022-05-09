import { ChangeDetectionStrategy, Component, OnInit } from '@angular/core';
import { StoreState } from '../../store/store.state';
import { Store } from '@ngrx/store';
import { getClaims, getPhoto } from 'src/app/store/store.actions';

declare const Office: any;
declare const Word: any;

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss'],
  changeDetection: ChangeDetectionStrategy.OnPush
})
export class HomeComponent implements OnInit {
  displayedColumns: string[] = ['claim', 'value'];
  datasource$ = this.store.select(s => s.store.claims);
  loggedIn$ = this.store.select(s => s.store.loggedIn);

  constructor(private store: Store<{ store: StoreState }>) { }

  ngOnInit(): void {
    // this.loggedIn$.subscribe((l) => {
    //   if (!l && Office.context.ui.messageParent) Office.context.ui.messageParent(JSON.stringify({ status: 'failed' }));
    // });
    this.store.dispatch(getClaims());
  }

  insert() {
    return Word.run((context: any) => {
      const paragraph = context.document.body.insertParagraph('Hello World', Word.InsertLocation.start);
      return context.sync();
    });
  }
}