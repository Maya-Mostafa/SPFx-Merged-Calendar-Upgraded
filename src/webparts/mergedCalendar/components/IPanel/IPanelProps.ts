import {IDropdownOption} from "@fluentui/react";

export interface IPanelProps{
    dpdOptions: IDropdownOption[];
    calSettings: {}[];
    isOpen: boolean;
    dismissPanel: any;
    isDataLoading: boolean;    
    showWeekends: boolean;
    onChkChange: (calSettings:{}) => ((ev: any, checked: boolean)=>void);
    onChkViewChange : (ev: any, checked: boolean) => void;
    onDpdChange: (calSettings:{}) => ((ev: any, item: IDropdownOption)=>void);
}