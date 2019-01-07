import { IWorkingWithWebPartProps } from "../WorkingWithWebPart";
import { DisplayMode } from '@microsoft/sp-core-library';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IWorkingWithProps extends IWorkingWithWebPartProps {
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  graphClient: MSGraphClient;
}
