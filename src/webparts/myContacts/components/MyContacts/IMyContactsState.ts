import { IContact } from '..';

export interface IMyContactsState {
  contacts: IContact[];
  error: string;
  loading: boolean;
}
