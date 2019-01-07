import { IMeeting } from '.';

export interface IMyCalendarState {
  error: string;
  loading: boolean;
  meetings: IMeeting[];
  renderedDateTime: Date;
}