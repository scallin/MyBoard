import { IPerson } from '.';

export interface IPersons {
    '@odata.context': string;
    value: IPerson[];
}