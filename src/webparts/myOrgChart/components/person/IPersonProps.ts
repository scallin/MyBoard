import { MSGraphClient } from '@microsoft/sp-http';
import { IPerson } from "..";

export interface IPersonProps {
  className: string;
  person: IPerson;
  graphClient: MSGraphClient;
}
