import * as React from 'react';
import './ILegend.scss';
import styles from '../MergedCalendar.module.scss';
import { ILegendProps } from './ILegendProps';

export default function ILegend(props:ILegendProps){
    
    return(
        <div className={styles.calendarLegend}>
            <ul>
            {
                props.calSettings.map((value:any)=>{
                    return(
                        <li key={value.Id}>
                            {(value.ShowCal ) && //&& value.CalType !== 'External'
                            <a href={value.LegendURL} target="_blank" data-interception="off">
                                <span className={styles.legendBullet +' calLegend_'+value.BgColor}></span>
                                    <span className={styles.legendText}>{value.Title}</span>
                            </a>
                            }
                        </li>
                    );
                })
            }
            </ul>
        </div>
    );
}