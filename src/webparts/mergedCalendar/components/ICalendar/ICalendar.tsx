import * as React from 'react';
import FullCalendar from '@fullcalendar/react';
import dayGridPlugin from '@fullcalendar/daygrid';
import timeGridPlugin from '@fullcalendar/timegrid';
import interactionPlugin from '@fullcalendar/interaction';
import rrulePlugin from '@fullcalendar/rrule';

import styles from '../MergedCalendar.module.scss';
import {ICalendarProps} from './ICalendarProps';

import {isUserManage} from '../../Services/WpProperties';

export default function ICalendar(props:ICalendarProps){

    return(
        <div className={styles.calendarCntnr}>
          <FullCalendar
            plugins = {
              [dayGridPlugin, timeGridPlugin, interactionPlugin, rrulePlugin]
            }
            headerToolbar = {{
              left: 'prev,next today',
              center: 'title',
              //right: isUserManage(props.context) ? 'dayGridMonth,timeGridWeek,timeGridDay settingsBtn addEventBtn' : 'dayGridMonth,timeGridWeek,timeGridDay addEventBtn' 
              right: isUserManage(props.context) ? 'dayGridMonth,timeGridWeek,timeGridDay settingsBtn' : 'dayGridMonth,timeGridWeek,timeGridDay' 
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
                
              }
            }}          
            eventTimeFormat={{
              hour: 'numeric',
              minute: '2-digit',
              meridiem: 'short'
            }}
            initialView='dayGridMonth'   
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
          />
      </div> 
    );
}