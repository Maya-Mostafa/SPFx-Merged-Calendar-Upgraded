import * as React from 'react';
import styles from '../MergedCalendar.module.scss';
import {IEventDetailsModernProps} from './IEventDetailsModernProps';
import * as moment from 'moment-timezone'; 

import {ActionButton} from '@fluentui/react';

export default function IEventDetailsModern (props: IEventDetailsModernProps){

    return(
        <div className={styles.eventDetailsModern}>
            
            <div className={styles.eventDateTitle}>
                <div className={styles.eventDate}>
                    <div className={styles.month}>{moment(new Date(props.EventCalDate)).format('MMM')}</div>
                    <div className={styles.day}>{moment(new Date(props.EventCalDate)).format('D')}</div>
                </div>
                <div>
                    <h5 className={styles.calendarTitle} style={{backgroundColor: props.CalendarColor, color: props.CalendarFontColor}}>{props.CalendarName}</h5>
                    <h2 className={styles.eventTitle}>{props.Title}</h2>
                </div>
            </div>

            <section>
                <label>When</label>
                <div>
                    {moment(new Date(props.EventCalDate)).format('dddd, MMMM Do YYYY')} <br/>
                    {moment(new Date(props.EventCalDate)).format('LT')} - {moment(new Date(props.EventCalEndDate)).format('LT')} <br/>
                    {props.AllDay &&
                        <i> (All Day Event)</i>
                    }
                </div>
            </section>

            <section>
                <label></label>
                <ActionButton iconProps={{ iconName: 'AddEvent' }} allowDisabledFocus onClick={()=>props.handleAddtoCal(props.Title, props.Body, props.EventCalDate, props.EventCalEndDate, props.Location)}>Add to my Calendar</ActionButton>
            </section>
            
            
            {props.Location &&
                <section>
                    <label>Where</label>
                    <div>
                        {props.Location}
                    </div>
                </section>
            }

            {props.Category &&
                <section>
                    <label>Category</label>
                    <div className={styles.evIp}>{props.Category}</div>
                </section>
            }

            {props.Body &&
                <div className={styles.eventBody}>
                    <label>About this event</label>
                    <div><p dangerouslySetInnerHTML={{__html: props.Body}}></p></div>
                </div>
            }
            

        </div>
    );
}