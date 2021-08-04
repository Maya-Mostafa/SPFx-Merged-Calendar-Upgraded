import {IDropdownOption} from "@fluentui/react";

export interface ICalSettingsProps{
    dpdOptions: IDropdownOption[];
    calSettings: {}[];
    showWeekends: boolean;
    onChkChange: (calSettings:{}) => ((ev: any, checked: boolean)=>void);
    onChkViewChange : (ev: any, checked: boolean) => void;
    onDpdChange: (calSettings:{}) => ((ev: any, item: IDropdownOption)=>void);
}