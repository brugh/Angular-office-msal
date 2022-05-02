import { Injectable } from '@angular/core';
import { Actions, createEffect, ofType } from '@ngrx/effects';
import { Store } from '@ngrx/store';
import { withLatestFrom, map } from 'rxjs';
import { GraphService } from '../services/graph.service';
import { getClaims, getPhoto, setClaims, setPhoto } from './store.actions';
import { StoreState } from './store.state';

@Injectable()
export class StoreEffects {
  client = this.graphService.getGraphClient();

  constructor(private store: Store<{ store: StoreState }>, private actions: Actions,
    private graphService: GraphService) { }

  public getClaims$ = createEffect(() => {
    return this.actions.pipe(
      ofType(getClaims),
      withLatestFrom(this.store),
      map(async () => {
        const data = await this.client.api('/me').get();
        this.store.dispatch(setClaims({
          claims: [
            { id: 1, claim: "Name", value: data ? data['givenName'] : null },
            { id: 2, claim: "Surname", value: data ? data['surname'] : null },
            { id: 3, claim: "User Principal Name (UPN)", value: data ? data['userPrincipalName'] : null },
            { id: 4, claim: "ID", value: data ? data['id'] : null },
          ]
        }));
      })
    )
  }, { dispatch: false});

  public getPhoto$ = createEffect(() => {
    return this.actions.pipe(
      ofType(getPhoto),
      withLatestFrom(this.store),
      map(async () => {
        const dataUrl = await this.client.api('/me/photo/$value').get();
        console.log(dataUrl);
        const reader = new FileReader();
        reader.onloadend = () => this.store.dispatch(setPhoto({ photo: reader.result }));
        reader.readAsDataURL(dataUrl);
      })
    )
  }, { dispatch: false});
}