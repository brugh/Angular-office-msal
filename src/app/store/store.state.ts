
export interface StoreState {
  token?: string | undefined;
  claims?: Claim[];
  loggedIn: boolean;
  photo?: ArrayBuffer | string | null;
}

export interface Claim {
  id: number,
  claim: string | undefined,
  value: string | undefined
}

export const initialState = {
  loggedIn: false,
};
