import * as React from 'react';
import { sliceEvents, createPlugin } from '@fullcalendar/core';
import styles from '../MergedCalendar.module.scss';
import * as moment from 'moment-timezone'; 

const CustomView = ({ start, end, events }) => (
  <div className={styles.peelUpcomingEvents}>
    <div>Custom View Testing 31 May</div>
    <div>
      {start.toUTCString()} to {end.toUTCString()}
    </div>
    <div>There are {events.length} events.</div>

    <ul>
      {events.map(event => {
        return(
          <li className={styles.eventItem}>
            <div className={styles.eventDate}>
                {event.def.extendedProps.recurr ?
                    <div className={styles.recurrentDate}>
                        <div>{moment(event.range.start).format('dd')} {moment(event.range.start).format('LLL')}</div>    
                        <div>{moment(event.range.end).format('dd')} {moment(event.range.end).format('LLL')}</div>
                    </div>
                    :
                    <>
                        <div className={styles.eventMonth}>{moment(event.range.start).format('MMM')}</div>
                        <div className={styles.eventDay}>{new Date(event.range.start).getDate()}</div>
                    </>
                }
            </div>
            <div className={styles.eventDetails}>
                {/* <h5><a onClick={props.eventClickHandler}>{event.def.title}</a></h5> */}
                <h5><a >{event.def.title}</a></h5>
                <div className={styles.eventTimes}>{moment(event.range.start).format('ff')} - {moment(event.range.end).format('tt')}</div>
                <div className={styles.eventLocation}>{event.def.extendedProps._location}</div>
            </div>
        </li>
        );
      })}
    </ul>

  </div>
);

const CustomViewPlugin = (props) => {
  const events = sliceEvents(props, true);
  const { dateProfile } = props;
  const { currentRange } = dateProfile;

  console.log("CustomViewPlugin props", props);
  console.log("CUSTOMVIEW events", events);

  return (
    <CustomView
      events={events}
      start={currentRange.start}
      end={currentRange.end}
    />
  );
};

export default createPlugin({
  name: 'custom',  
  views: {
    custom: {
      content: CustomViewPlugin,
      duration: { days: 7 },
      type: 'listWeek',
      visibleRange: {start: new Date().setDate(new Date().getDate()), end: new Date().setDate(new Date().getDate()+7)}
    },
  },
    
    
});