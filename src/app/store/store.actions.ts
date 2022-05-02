import { createAction, props } from '@ngrx/store';
import { Claim } from './store.state';

export const setLoggedIn = createAction('[Token] Login', props<{loggedIn: boolean}>());
export const storeToken = createAction('[Token] Store', props<{token: string}>());
export const storeReset = createAction('[Token] Reset');
export const getClaims = createAction('[Profile] Get claims');
export const setClaims = createAction('[Profile] Store claims', props<{claims: Claim[]}>());
export const getPhoto = createAction('[Profile] Get photo');
export const setPhoto = createAction('[Profile] Store photo', props<{photo: ArrayBuffer | string | null}>());
export const nullAction = createAction('[null]');