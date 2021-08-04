import { WebPartContext } from "@microsoft/sp-webpart-base";
import {IDropdownOption} from "@fluentui/react";

export interface IMergedCalendarProps {
  description: string;
  showWeekends: boolean;
  context: WebPartContext;  
  calSettingsList: string;
  dpdOptions: IDropdownOption[];
}
