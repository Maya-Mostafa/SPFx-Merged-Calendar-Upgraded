import * as React from 'react';
import './ILegend.scss';
import {Checkbox, Link} from '@fluentui/react';
import styles from '../MergedCalendar.module.scss';
import { ILegendProps } from './ILegendProps';
import { initializeIcons } from '@uifabric/icons';
// import { getPosGrpMapping} from '../../Services/CalendarRequests';

export default function ILegend(props:ILegendProps){    
    initializeIcons();
    const _renderLabelWithLink = (calTitle: string, hrefVal: string) => {
        return (
            <Link href={hrefVal} target="_blank" underline>
                {calTitle}
            </Link>
        );
    };

    {/* <a href={value.LegendURL} target="_blank" data-interception="off">
        <span className={styles.legendBullet +' calLegend_'+value.BgColor}></span>
        <span className={styles.legendText}>{value.Title}</span>
    </a> */}

    const isUserGrpCal = (calView: string) => {
        // console.log("calView", calView);
        if (props.posGrps[calView.trim()] == undefined) return true;
        else{
            for (let userGrp of props.userGrps){
                if (props.posGrps[calView.trim()] && props.posGrps[calView.trim()].indexOf(Number(userGrp)) !== -1){
                    return true;
                }
            }
            return false;
        }
        
    };
    
    // console.log("props.calSettings", props.calSettings);
    // console.log("legend userGrps", props.userGrps);
    // console.log("props.posGrps", props.posGrps);

    const sortedCalSettings = props.calSettings.sort((a,b) => a.Title.localeCompare(b.Title));

    return(
        <div className={styles.calendarLegend}>
            <ul>
            {
                sortedCalSettings.map((value:any)=>{
                    return(
                        <React.Fragment>
                            {value.ShowCal  && //&& value.CalType !== 'External'
                                <li key={value.Id}>    
                                    <Checkbox 
                                        className={'chkboxLegend chkbox_'+value.BgColor}
                                        label={value.Title} 
                                        defaultChecked={isUserGrpCal(value.View)}
                                        // checked={props.legendChked}
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