import { createReducer, on } from '@ngrx/store';
import { storeToken, storeReset, setLoggedIn, setClaims, setPhoto } from './store.actions';
import { initialState, StoreState } from './store.state';

export const storeReducer = createReducer<StoreState>(
  initialState,
  on(storeReset, () => { return { ...initialState }; }),
  on(setLoggedIn, (state: StoreState, action): StoreState => { 
    console.log("Reducing ", state, action.loggedIn);
    return { ...state, loggedIn: action.loggedIn }; 
  }),
  on(storeToken, (state, action) => { return { ...state, token: action.token }; }),
  on(setClaims, (state, action) => { return { ...state, claims: action.claims }; }),
  on(setPhoto, (state, action) => { console.log(action.photo); return { ...state, photo: action.photo }; })
);
