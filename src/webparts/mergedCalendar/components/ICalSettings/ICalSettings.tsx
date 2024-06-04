import * as React from 'react';
import styles from '../MergedCalendar.module.scss';
import { ICalSettingsProps } from './ICalSettingsProps';
import { initializeIcons } from '@uifabric/icons';
import { Icon, Checkbox, Stack, Dropdown, Label } from 'office-ui-fabric-react';

export default function ICalSettings (props:ICalSettingsProps) {

    const stackTokens = { childrenGap: 20 , maxWidth: 250};    
    initializeIcons();
    const CalendarIcon = () => <Icon iconName="Calendar" />;
    const CalendarSettingsIcon = () => <Icon iconName="CalendarSettings" />;

    return(
        <div className={styles.calendarPanel}>
            <h3 className={styles.panelHeader}><Label><CalendarIcon />Choose Calendar(s)</Label></h3>
            <Stack tokens={stackTokens}>
                {props.calSettings.map((value:any) => {        
                    return (
                        <div>                            
                            <Checkbox key={value.Id} onChange={props.onChkChange(value)} defaultChecked={value.Chkd} label={value.Title} disabled={value.Disabled} />
                            {value.Dpd &&
                                <Dropdown onChange={props.onDpdChange(value)} className={styles.marginT5} placeholder="Select Day..."  defaultSelectedKey={value.CalName} options={props.dpdOptions} />
                            }
                        </div>
                    );
                })}
            </Stack>
            <hr className={styles.panelSeparator}/>
            <h3 className={styles.panelHeader}><Label><CalendarSettingsIcon />Settings</Label></h3>
            <Checkbox label="Show Weekends" defaultChecked={props.showWeekends} onChange={props.onChkViewChange}/>
        </div>
    );
}