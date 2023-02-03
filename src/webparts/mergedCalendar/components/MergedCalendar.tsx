import * as React from 'react';
import styles from './MergedCalendar.module.scss';
import { IMergedCalendarProps } from './IMergedCalendarProps';
//import { escape } from '@microsoft/sp-lodash-subset';

import {IDropdownOption, MessageBar, MessageBarType} from '@fluentui/react';
import {useBoolean} from '@fluentui/react-hooks';

import {CalendarOperations} from '../Services/CalendarOperations';
import {getCalSettings, updateCalSettings} from '../Services/CalendarSettingsOps';
import {addToMyGraphCal, getMySchoolCalGUID, reRenderCalendars, calsErrs, getUserGrp, getAllPosGrps, getLegendChksState} from '../Services/CalendarRequests';
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
  const [calMsgErrs, setCalMsgErrs] = React.useState([]);
  const [userGrps, setUserGrps] = React.useState([]);
  const [posGrps, setPosGrps] = React.useState([]);

  const [calVisibility, setCalVisibility] = React.useState <{calId: string, calChk: boolean}>({calId: null, calChk: null});
  const [calsVisibility, setCalsVisibility] = React.useState([]);

  const calSettingsList = props.calSettingsList ? props.calSettingsList : "CalendarSettings";
  const legendPos = props.legendPos ? props.legendPos : "top";
  const legendAlign = props.legendAlign ? props.legendAlign : "horizontal";

  const spCalParams = props.spCalParams ? props.spCalParams : {rangeStart: 3, rangeEnd: 4, pageSize: 150};
  const graphCalParams = props.graphCalParams ? props.graphCalParams :{rangeStart: '3', rangeEnd: '4', pageSize: '150'};

  const [currentCalDate, setCurrentCalDate] = React.useState(new Date().toISOString());

  // const calSettingsList = props.calSettingsList ;
  React.useEffect(()=>{
    getUserGrp(props.context).then(userGrpsResult => {
      setUserGrps(userGrpsResult);
      // console.log("userGrpsResult", userGrpsResult);
      getAllPosGrps(props.context).then(posGrpsResult => {
        setPosGrps(posGrpsResult);
        
        _calendarOps.displayCalendars(props.context, calSettingsList, currentCalDate, userGrpsResult, posGrpsResult, Number(props.spCalPageSize), graphCalParams).then((result:{}[])=>{
          //setEventSources(result);
          //setEventSources(reRenderCalendars(result, calVisibility));
          //console.log("setEventSources", result);
          setCalMsgErrs(calsErrs);
          if (calsVisibility.length > 1){
            setEventSources(prevEventSources => {
              let tempEventSources = [];
              for (let calVis of calsVisibility){
                tempEventSources = reRenderCalendars(prevEventSources, calVis);
              }
              return [...tempEventSources];
            });
          }else{
            setEventSources(result);
          }
        });

        getCalSettings(props.context, calSettingsList).then((result:{}[])=>{
          setCalSettings(result);
        });
        
      });
    });

    getMySchoolCalGUID(props.context, calSettingsList).then((result)=>{
      setListGUID(result);
    }); 

  },[currentCalDate]);

  React.useEffect(()=>{
    console.log("useEffect calVisibility -->", calVisibility);
    setEventSources(prevEventSources => reRenderCalendars(prevEventSources, calVisibility));
    setCalsVisibility(prevCalsVisibility => getLegendChksState(prevCalsVisibility, calVisibility));
  },[calVisibility]);

  const chkHandleChange = (newCalSettings:{})=>{    
    return (ev: any, checked: boolean) => { 
      toggleIsDataLoading();
      updateCalSettings(props.context, calSettingsList, newCalSettings, checked).then(()=>{
        _calendarOps.displayCalendars(props.context, calSettingsList, currentCalDate, userGrps, posGrps, Number(props.spCalPageSize), graphCalParams).then((result:{}[])=>{
          setEventSources(result);
          toggleIsDataLoading();
        });
        getCalSettings(props.context, calSettingsList).then((result:{}[])=>{
          setCalSettings(result);
        });
      });
      
     };
  };  
  const dpdHandleChange = (newCalSettings:any)=>{
    return (ev: any, item: IDropdownOption) => { 
      toggleIsDataLoading();
      updateCalSettings(props.context, calSettingsList, newCalSettings, newCalSettings.ShowCal, item.key).then(()=>{
        _calendarOps.displayCalendars(props.context, calSettingsList, currentCalDate, userGrps, posGrps, Number(props.spCalPageSize), graphCalParams).then((result:{}[])=>{
          setEventSources(result);
          toggleIsDataLoading();
        });
        getCalSettings(props.context, calSettingsList).then((result:{}[])=>{
          setCalSettings(result);
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
    //console.log("ev details arg", arg);
    //console.log(formatEvDetails(arg));
    setEventDetails(formatEvDetails(arg));
    toggleHideDialog();
  };

  const handleAddtoCal = ()=>{
    addToMyGraphCal(props.context).then((result)=>{
      // console.log('calendar updated', result);
    });
  };

  const onLegendChkChange = (calId: string) =>{
    return(ev: any, checked: boolean) =>{
      // setCalVisibility({calId: calId, calChk: checked});
      //console.log("onLegendChkChange --> calId+checked.toString()", calId+checked.toString());
      const newCalVisibility = {calId: calId, calChk: checked};
      setCalVisibility({...newCalVisibility});
    };
  };

  const passCurrentDate = (currDate: string) => {
    console.log("passCurrentCalDate function", currDate);
    setCurrentCalDate(currDate);
  };

  return(
    <div className={styles.mergedCalendar}>

      {legendPos === 'top' &&
        <div className={`${styles.legendTop} ${legendAlign === 'horizontal' ? styles.legendHz : '' }`}>
          <ILegend
            calSettings={calSettings} 
            onLegendChkChange={onLegendChkChange}
            legendChked = {true}
            userGrps = {userGrps}
            posGrps = {posGrps}
          />
        </div>
      }

      <ICalendar 
        eventSources={eventSources} 
        // showWeekends={props.showWeekends ? props.showWeekends : false } 
        showWeekends={showWeekends}
        calSettings={calSettings}
        openPanel={openPanel}
        handleDateClick={handleDateClick}
        context={props.context}
        listGUID = {listGUID}
        passCurrentDate = {passCurrentDate}
      />

      <IPanel
        dpdOptions={props.dpdOptions} 
        calSettings={calSettings}
        onChkChange={chkHandleChange}
        onDpdChange={dpdHandleChange}
        isOpen = {isOpen}
        dismissPanel = {dismissPanel}
        isDataLoading = {isDataLoading} 
        showWeekends= {showWeekends} 
        onChkViewChange= {chkViewHandleChange}
      />

      {legendPos === 'bottom' &&
        <div className={legendAlign === 'horizontal' ? styles.legendHz : '' }>
          <ILegend 
            calSettings={calSettings} 
            onLegendChkChange={onLegendChkChange}
            legendChked = {true}
            userGrps = {userGrps}
            posGrps = {posGrps}
          />
        </div>
      }

      <IDialog 
        hideDialog={hideDialog} 
        toggleHideDialog={toggleHideDialog}
        eventDetails={eventDetails}
        handleAddtoCal = {handleAddtoCal}
      />
      
      {calMsgErrs.length > 0 &&
        <MessageBar className={styles.calErrsMsg} messageBarType={MessageBarType.warning}>
          Warning! Calendar Errors, please check
          <ul>
            {calMsgErrs.map((msg)=>{
              return <li>{msg}</li> ;
            })}
          </ul>
        </MessageBar>
      }
    </div>
  );
  
  
}
