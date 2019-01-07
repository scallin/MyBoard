import { IContact } from '.';

export interface IWorkingWithState {
  recentContacts: IContact[];
  error: string;
  loading: boolean;
}
