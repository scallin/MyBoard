import { DisplayMode } from '@microsoft/sp-core-library';
import { IMyContactsWebPartProps } from "../../MyContactsWebPart";
import { MSGraphClient } from '@microsoft/sp-http';

export interface IMyContactsProps extends IMyContactsWebPartProps{
  displayMode: DisplayMode;
  graphClient: MSGraphClient;
  updateProperty: (value: string) => void;
}
