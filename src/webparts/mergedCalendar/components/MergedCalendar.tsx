import * as React from 'react';
import styles from './MergedCalendar.module.scss';
import { IMergedCalendarProps } from './IMergedCalendarProps';
//import { escape } from '@microsoft/sp-lodash-subset';

import {IDropdownOption} from '@fluentui/react';
import {useBoolean} from '@fluentui/react-hooks';

import {CalendarOperations} from '../Services/CalendarOperations';
import {getCalSettings, updateCalSettings} from '../Services/CalendarSettingsOps';
import {addToMyGraphCal, getMySchoolCalGUID} from '../Services/CalendarRequests';
import {formatEvDetails} from '../Services/EventFormat';
import {setWpData} from '../Services/WpProperties';

import ICalendar from './ICalendar/ICalendar';
import IPanel from './IPanel/IPanel';
import ILegend from './ILegend/ILegend';
import IDialog from './IDialog/IDialog';

export default function MergedCalendar (props:IMergedCalendarProps) {
  
  const _calendarOps = new CalendarOperations();
  const [eventSources, setEventSources] = React.useState([]);
  const [calSettings, setCalSettings] = React.useState([]);
  const [isOpen, { setTrue: openPanel, setFalse: dismissPanel }] = useBoolean(false);
  const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true);
  const [eventDetails, setEventDetails] = React.useState({});
  const [isDataLoading, { toggle: toggleIsDataLoading }] = useBoolean(false);
  const [showWeekends, { toggle: toggleshowWeekends }] = useBoolean(props.showWeekends);
  const [listGUID, setListGUID] = React.useState('');

  const calSettingsList = props.calSettingsList ? props.calSettingsList : "CalendarSettings";
  // const calSettingsList = props.calSettingsList ;
  React.useEffect(()=>{
    _calendarOps.displayCalendars(props.context, calSettingsList).then((result:{}[])=>{
      setEventSources(result);
    });
    getCalSettings(props.context, calSettingsList).then((result:{}[])=>{
      setCalSettings(result);
    });
    getMySchoolCalGUID(props.context, calSettingsList).then((result)=>{
      setListGUID(result);
    });
  },[eventSources.length]);

  const chkHandleChange = (newCalSettings:{})=>{    
    return (ev: any, checked: boolean) => { 
      toggleIsDataLoading();
      updateCalSettings(props.context, calSettingsList, newCalSettings, checked).then(()=>{
        _calendarOps.displayCalendars(props.context, calSettingsList).then((result:{}[])=>{
          setEventSources(result);
          toggleIsDataLoading();
        });
      });
     };
  };  
  const dpdHandleChange = (newCalSettings:any)=>{
    return (ev: any, item: IDropdownOption) => { 
      toggleIsDataLoading();
      updateCalSettings(props.context, props.calSettingsList, newCalSettings, newCalSettings.ShowCal, item.key).then(()=>{
        _calendarOps.displayCalendars(props.context, props.calSettingsList).then((result:{}[])=>{
          setEventSources(result);
          toggleIsDataLoading();
        });
      });
     };
  };
  const chkViewHandleChange = (ev: any, checked: boolean) =>{
    toggleIsDataLoading();
    setWpData(props.context, "showWeekends", checked).then(()=>{
      toggleshowWeekends();
      toggleIsDataLoading();
    });
    
  };
  const handleDateClick = (arg:any) =>{
    //console.log(arg);
    //console.log(formatEvDetails(arg));
    setEventDetails(formatEvDetails(arg));
    toggleHideDialog();
  };

  const handleAddtoCal = ()=>{
    addToMyGraphCal(props.context).then((result)=>{
      console.log('calendar updated', result);
    });
  };

  return(
    <div className={styles.mergedCalendar}>

      <ICalendar 
        eventSources={eventSources} 
        // showWeekends={props.showWeekends ? props.showWeekends : false } 
        showWeekends={showWeekends}
        calSettings={calSettings}
        openPanel={openPanel}
        handleDateClick={handleDateClick}
        context={props.context}
        listGUID = {listGUID}/>

      <IPanel
        dpdOptions={props.dpdOptions} 
        calSettings={calSettings}
        onChkChange={chkHandleChange}
        onDpdChange={dpdHandleChange}
        isOpen = {isOpen}
        dismissPanel = {dismissPanel}
        isDataLoading = {isDataLoading} 
        showWeekends= {showWeekends} 
        onChkViewChange= {chkViewHandleChange}/>

      <ILegend calSettings={calSettings} />

      <IDialog 
        hideDialog={hideDialog} 
        toggleHideDialog={toggleHideDialog}
        eventDetails={eventDetails}
        handleAddtoCal = {handleAddtoCal}
        />
    
    </div>
  );
  
  
}
