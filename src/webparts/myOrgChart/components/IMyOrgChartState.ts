import { IPerson } from '.';

export interface IMyOrgChartState {
  manager: IPerson;
  loadingMgr: boolean;
  errorMgr: string;
  user: IPerson;
  loadingUser: boolean;
  errorUser: string;
  reports: IPerson[];
  loadingReps: boolean;
  errorReps: string;
}
