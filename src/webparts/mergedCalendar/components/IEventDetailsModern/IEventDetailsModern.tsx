import * as React from 'react';
import styles from '../MergedCalendar.module.scss';
import {IEventDetailsModernProps} from './IEventDetailsModernProps';
import * as moment from 'moment-timezone'; 
import {ActionButton} from 'office-ui-fabric-react';

export default function IEventDetailsModern (props: IEventDetailsModernProps){

    const [eventAdded, setEventAdded] = React.useState(props.EventAdded);
    const addToMyCalHandler = () => {
        setEventAdded(true);
        props.handleAddtoCal(props.Title, props.Body, props.AllDay ? props.Start : props.EventCalDate, props.AllDay ? props.End : props.EventCalEndDate, props.Location);
    };

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
                    
                    {props.AllDay 
                        ?
                        <div>
                            <div>
                                {moment(new Date(props.EventCalDate)).format('dddd, MMMM Do YYYY')}
                                {/* {moment(new Date(props.Start)).format('dddd, MMMM Do YYYY')} */}
                            </div>
                            <i>(All Day Event)</i>
                        </div>
                        :
                        <div>
                            {moment(new Date(props.EventCalDate)).format('dddd, MMMM Do YYYY')} <br/>
                            {moment(new Date(props.EventCalDate)).format('LT')} - {moment(new Date(props.EventCalEndDate)).format('LT')} <br/>
                        </div>
                    }
                </div>
            </section>
            
            {(props.showAddToCal === true || props.showAddToCal === undefined) &&
                <section>
                    <label></label>
                    {eventAdded 
                        ?
                        <ActionButton iconProps={{ iconName: 'EventAccepted' }} allowDisabledFocus disabled>Added to my Calendar</ActionButton>
                        :
                        <ActionButton iconProps={{ iconName: 'AddEvent' }} allowDisabledFocus onClick={addToMyCalHandler}>Add to my Calendar</ActionButton>
                    }
                </section>
            }
            
            
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