import { createReducer, on } from '@ngrx/store';
import { storeToken, storeReset, setLoggedIn, setClaims, setPhoto } from './store.actions';
import { initialState, StoreState } from './store.state';

export const storeReducer = createReducer<StoreState>(
  initialState,
  on(storeReset, () => { return { ...initialState }; }),
  on(setLoggedIn, (state: StoreState, { loggedIn }): StoreState => { return { ...state, loggedIn: loggedIn }; }),
  on(setClaims, (state: StoreState, { claims }): StoreState => { return { ...state, claims: claims }; }),
  on(setPhoto, (state: StoreState, { photo }): StoreState => { return { ...state, photo: photo }; }),
  on(storeToken, (state: StoreState, { token }): StoreState => { return { ...state, token: token }; }),
);
