import { WebPartContext } from "@microsoft/sp-webpart-base";
import {IDropdownOption} from "@fluentui/react";

export interface IMergedCalendarProps {
  description: string;
  showWeekends: boolean;
  context: WebPartContext;  
  calSettingsList: string;
  dpdOptions: IDropdownOption[];
  legendPos: string;
  legendAlign: string;
  spCalParams : {rangeStart: number, rangeEnd: number, pageSize: number};
  graphCalParams : {rangeStart: number, rangeEnd: number, pageSize: number};
  spCalPageSize: number;
}
