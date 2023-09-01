import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from '@microsoft/sp-core-library';
export interface IDelphiBirthsProps {
  title: string;
  numberUpcomingDays: number;
  context: WebPartContext;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  imageTemplate:string;
  height?:string;
  width?:string;
}
