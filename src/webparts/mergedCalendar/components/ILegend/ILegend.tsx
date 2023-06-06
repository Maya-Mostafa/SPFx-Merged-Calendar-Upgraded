import * as React from 'react';
import './ILegend.scss';
import {Checkbox, Link} from '@fluentui/react';
import styles from '../MergedCalendar.module.scss';
import { ILegendProps } from './ILegendProps';
import { initializeIcons } from '@uifabric/icons';

export default function ILegend(props:ILegendProps){    
    initializeIcons();
    const _renderLabelWithLink = (calTitle: string, hrefVal: string) => {
        return (
            <Link href={hrefVal} target="_blank" underline>
                {calTitle}
            </Link>
        );
    };
    const sortedCalSettings = props.calSettings.sort((a,b) => a.Title.localeCompare(b.Title));
    const isItemChkd = (items: any, itemId: any) => {
        return items.find(item => item.calId === itemId).calChk;
    };

    //console.log("props.calSettings", props.calSettings);
    //console.log("props.legendChked", props.legendChked);
    return(
        <div className={styles.calendarLegend}>
            <ul>
                <li>
                    <Checkbox 
                        label='All' 
                        checked = {props.legendChked.length > 1 && isItemChkd(props.legendChked, 'all')} 
                        onChange={props.onLegendChkChange('all')} 
                    />
                </li>
                {
                    sortedCalSettings.map((value:any)=>{
                        return(
                            <React.Fragment>
                                {value.ShowCal  && //&& value.CalType !== 'External'
                                    <li key={value.Id} id={`legend-item-${value.Id}`}>    
                                        <Checkbox 
                                            className={'chkboxLegend chkbox_'+value.BgColor}
                                            label={value.Title} 
                                            checked={props.legendChked.length > 1 && isItemChkd(props.legendChked, value.Id)}
                                            onChange={props.onLegendChkChange(value.Id)} 
                                            onRenderLabel={() => _renderLabelWithLink(value.Title, value.LegendURL)}
                                        />
                                    </li>
                                }
                            </React.Fragment>
                        );
                    })
                }
            </ul>
        </div>
    );
}