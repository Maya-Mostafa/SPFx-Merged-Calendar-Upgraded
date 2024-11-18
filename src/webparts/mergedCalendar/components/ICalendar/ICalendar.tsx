import * as React from 'react';
import FullCalendar from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import listPlugin from '@fullcalendar/list';
import interactionPlugin from '@fullcalendar/interaction';
import rrulePlugin from '@fullcalendar/rrule';

import { initializeIcons } from '@uifabric/icons';
import styles from '../MergedCalendar.module.scss';
import '../MergedCalendar.scss';
import {ICalendarProps} from './ICalendarProps';
import {isUserManage} from '../../Services/WpProperties';
import CustomViewPlugin from './CustomViewPlugin';

export default function ICalendar(props:ICalendarProps){

  console.log("ICalendarProps", props);

  initializeIcons();

  const calendarRef = React.useRef<any>();
  const calendarNext = () => {
    let calendarApi = calendarRef.current.getApi();
    calendarApi.next();
  };
  const calendarPrev = () => {
    let calendarApi = calendarRef.current.getApi();
    calendarApi.prev();
  };

  let leftHdrButtons = 'customPrev,customNext today';
  let centerButtons = 'title';
  let rightButtons = isUserManage(props.context) ? 'dayGridMonth,timeGridWeek,timeGridDay,listMonth settingsBtn' : 'dayGridMonth,timeGridWeek,timeGridDay,listMonth';
  
  // if (props.isListView){
  //   leftHdrButtons = props.listViewNavBtns ? 'customPrev,customNext today' : '' ;
  //   centerButtons = props.listViewMonthTitle ? 'title' : '';
  //   if (isUserManage(props.context)) rightButtons = props.listViewViews ? 'dayGridMonth,timeGridWeek,timeGridDay,listMonth settingsBtn' : '';
  //   else rightButtons = props.listViewViews ? 'dayGridMonth,timeGridWeek,timeGridDay' : '';
  // }

  return(
      <div className={styles.calendarCntnr}>
        <FullCalendar
          ref={calendarRef}
          contentHeight = {props.isListView ? props.listViewHeight : 'auto'}
          plugins = {
            [dayGridPlugin, timeGridPlugin, interactionPlugin, rrulePlugin, listPlugin, CustomViewPlugin]
          }
          headerToolbar = {{
            //right: isUserManage(props.context) ? 'dayGridMonth,timeGridWeek,timeGridDay settingsBtn addEventBtn' : 'dayGridMonth,timeGridWeek,timeGridDay addEventBtn' 
            // left: 'prev,next today customPrev customNext',
            left: props.listViewNavBtns === false ? '' : leftHdrButtons,
            center: props.listViewMonthTitle === false ? '' : centerButtons,
            right: props.listViewViews === false ? '' : rightButtons 
          }}
          customButtons = {{
            settingsBtn : {
              text : 'Settings',
              click : props.openPanel,
            },
            addEventBtn : {
              text: 'Add Event',
              click : function(){
                window.open(
                  props.context.pageContext.web.absoluteUrl + '/_layouts/15/Event.aspx?ListGuid='+ props.listGUID +'&Mode=Edit',
                  '_blank' 
                );
              }
            },
            customPrev: {
              icon: 'chevron-left',
              click: function() {
                props.passCurrentDate(calendarRef.current.getApi().getDate().toISOString());
                calendarPrev();
              }
            },
            customNext: {
              icon:'chevron-right',
              click: function() {
                props.passCurrentDate(calendarRef.current.getApi().getDate().toISOString());
                calendarNext();
              }
            }
          }}          
          eventTimeFormat={{
            hour: 'numeric',
            minute: '2-digit',
            meridiem: 'short'
          }}
          // initialView = {props.isListView ? props.listViewType : 'dayGridMonth'} 
          eventClassNames={styles.eventItem}           
          editable={false}
          selectable={true}
          selectMirror={true}
          dayMaxEvents={false}
          displayEventEnd={true}
          eventDisplay='block'
          weekends={props.showWeekends}
          eventClick={props.handleDateClick}
          eventSources = {props.eventSources}
          views={{
            upcomingEventsGrid: {
              type: 'listMonth',
              duration: { days: props.viewDuration ? Number(props.viewDuration) : 7 },
            },
            upcomingEventsBox: {
              type: 'listMonth',
              duration: { days: props.viewDuration ? Number(props.viewDuration) : 7 },
              viewClassNames: 'peelUpcomingEventsView'
            }
          }}
          initialView = {props.calendarView ? props.calendarView : 'dayGridMonth'}    
          // visibleRange = {(currentDate) => {
          //   // Generate a new date for manipulating in the next step
          //   const startDate = new Date(currentDate.valueOf());
          //   const endDate = new Date(currentDate.valueOf());
        
          //   // Adjust the start & end dates, respectively
          //   startDate.setDate(startDate.getDate()); // One day in the past
          //   endDate.setDate(endDate.getDate() + 7); // Two days into the future

          //   console.log("startDate", startDate);
          //   console.log("endDate", endDate);

          //   return { start: startDate, end: endDate };
          // }}
          visibleRange={{
            start: new Date().setDate(new Date().getDate()), 
            end: new Date().setDate(new Date().getDate()+props.viewDuration)
          }}
          eventContent = {(eventInfo)=>{
            // console.log("eventInfo", eventInfo);
            return (
                <div>
                  {/* <div><b>{eventInfo.event._def.extendedProps._startTime} - {eventInfo.event._def.extendedProps._endTime}</b></div> */} 
                  <b>{eventInfo.timeText && eventInfo.timeText + ' '}</b>
                  <i>{eventInfo.event.title}</i>
                </div>
            );
          }}
        />
    </div> 
  );
}