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
  spCalParams : {rangeStart: string, rangeEnd: string, pageSize: string};
  graphCalParams : {rangeStart: string, rangeEnd: string, pageSize: string};
  spCalPageSize: string;

  isListView: boolean;
  listViewType: any;
  listViewNavBtns: boolean;
  listViewLegend: boolean;
  listViewErrors: boolean;
  listViewMonthTitle: boolean;
  listViewViews: boolean;
  listViewHeight: number;
  listViewTitle: string;

  posGrpView: boolean;
  calendarView: string;
  viewDuration: number;
}
