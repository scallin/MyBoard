import { IMyOrgChartWebPartProps } from "../MyOrgChartWebPart";
import { DisplayMode } from '@microsoft/sp-core-library';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IMyOrgChartProps extends IMyOrgChartWebPartProps {
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  graphClient: MSGraphClient;
}
