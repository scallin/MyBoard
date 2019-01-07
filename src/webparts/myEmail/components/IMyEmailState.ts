import { IMessage } from '.';

export interface IMyEmailState {
  error: string;
  loading: boolean;
  messages: IMessage[];
}