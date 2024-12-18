import * as React from 'react';
import styles from './MergedCalendar.module.scss';
import { IMergedCalendarProps } from './IMergedCalendarProps';
//import { escape } from '@microsoft/sp-lodash-subset';

import {IDropdownOption, MessageBar, MessageBarType, Label} from 'office-ui-fabric-react';
// import {useBoolean} from '@fluentui/react-hooks';
import { useBoolean } from '@uifabric/react-hooks';

import {CalendarOperations} from '../Services/CalendarOperations';
import {getCalSettings, isPosGrpsCal, isUserGrpCal, updateCalSettings} from '../Services/CalendarSettingsOps';
import {addToMyGraphCal, getMySchoolCalGUID, reRenderCalendars, calsErrs, getUserGrp, getAllPosGrps, getLegendChksState, getRotaryCals, getMembershipGroups, getListItemsGraph} from '../Services/CalendarRequests';
import {formatEvDetails} from '../Services/EventFormat';
import {setWpData} from '../Services/WpProperties';

import ICalendar from './ICalendar/ICalendar';
import IPanel from './IPanel/IPanel';
import ILegend from './ILegend/ILegend';
import IDialog from './IDialog/IDialog';

export default function MergedCalendar (props:IMergedCalendarProps) {  
  
  //const viewType = 'upcoming events';
  
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

  const [calsVisibility, setCalsVisibility] = React.useState([]);
  const [rotaryCals, setRotaryCals] = React.useState([]);

  const calSettingsList = props.calSettingsList ? props.calSettingsList : "CalendarSettings";
  const legendPos = props.legendPos ? props.legendPos : "top";
  const legendAlign = props.legendAlign ? props.legendAlign : "horizontal";

  const spCalParams = props.spCalParams ? props.spCalParams : {rangeStart: 3, rangeEnd: 4, pageSize: 150};
  const graphCalParams = props.graphCalParams ? props.graphCalParams :{rangeStart: '3', rangeEnd: '4', pageSize: '150'};

  const [currentCalDate, setCurrentCalDate] = React.useState(new Date().toISOString());

  const [clkdEvId, setClkdEvId] = React.useState(null);
  const [addedToMyCal, setAddedToMyCal] = React.useState([]);

  // let showErrors = true;
  // let showLegend = true;
  // let isListView = props.calendarView !== 'dayGridMonth';
  
  // if (isListView){ //old
  //   showErrors = props.listViewErrors;
  //   showLegend = props.listViewLegend;
  // }

  // if (props.listViewLegend !== undefined) showLegend = props.listViewLegend;
  // if (props.listViewErrors !== undefined) showErrors = props.listViewErrors;

  // if (props.calendarView === undefined)  {
  //   props.calendarView = "dayGridMonth";
  //   showLegend = true;
  //   showErrors = true;
  // }  

  let isListView = props.calendarView !== 'dayGridMonth';
  let showErrors = props.listViewErrors === false ? false : true;
  let showLegend = props.listViewLegend === false ? false : true;
  
  // reading the graph rotart calendars
  React.useEffect(()=>{

    // getListItemsGraph(props.context, "https://pdsb1.sharepoint.com/sites/ModernDemos", "Activities").then(res => {
    //   console.log("------------GRAPH API EVENTS", res);
    // })

    // getMembershipGroups(props.context).then(res => {
    //   console.log("Membership Groups", res);
    // });

    getRotaryCals(props.context).then(res =>{
      setRotaryCals(res);
    });
  }, []);

  // const calSettingsList = props.calSettingsList ;
  React.useEffect(()=>{

    console.log("calSettingsList", calSettingsList);
    console.log("Merged Calendar Props", props);

    getUserGrp(props.context).then(userGrpsResult => {
      setUserGrps(userGrpsResult);
      // console.log("userGrpsResult", userGrpsResult);
      getAllPosGrps(props.context).then(posGrpsResult => {
        let posGrpsResultMod = posGrpsResult;
        if (props.calendarView !== 'dayGridMonth' && !props.listViewLegend){          
          posGrpsResultMod = null;
          setPosGrps(null);
        }          
        else setPosGrps(posGrpsResult);
        
        _calendarOps.displayCalendars(props.context, calSettingsList, currentCalDate, userGrpsResult, posGrpsResultMod, Number(props.spCalPageSize), graphCalParams).then((result:{}[])=>{
          //setEventSources(result);
          console.log("setEventSources", result);
          setCalMsgErrs(calsErrs);

          if (calsVisibility.length > 1){
            // setEventSources(prevEventSources => reRenderCalendars(prevEventSources, calsVisibility));
            setEventSources(reRenderCalendars(result, calsVisibility));
          }else{
            setEventSources(result);
          }

        });

        getCalSettings(props.context, calSettingsList).then((result:any)=>{
          setCalSettings(result);
          
          // setting the legend checboxes visibility
          if (calsVisibility.length === 0){ // on first load
            const legend =  result.map(calItem => {
              return {
                calId: calItem.Id,
                calChk: isUserGrpCal(calItem.View, posGrpsResult, userGrpsResult),
                calRender: calItem.Chkd
              };
            });
            const renderedCalsLen = legend.filter((item: any) => item.calRender).length;
            const chkdCalsLen = legend.filter((item: any) => item.calChk && item.calRender).length;
            setCalsVisibility([{calId: 'all', calChk: renderedCalsLen === chkdCalsLen, calRender: true}, ...legend]);
          }else{ // on next & prev month
            setCalsVisibility(prev => {
              const clonePrev = [...prev];
              const renderedCalsLen = clonePrev.filter((item: any) => item.calRender && item.calId !== 'all').length;
              const chkdCalsLen = clonePrev.filter((item: any) => item.calChk && item.calRender && item.calId !== 'all').length;
              return clonePrev.map(item => {
                if (item.calId === 'all') item.calChk = renderedCalsLen === chkdCalsLen;
                return item;
              });
            });
          }
          
        });
        
      });
    });

    // used for adding event to my school calendar - disabled for now as it is not used
    // getMySchoolCalGUID(props.context, calSettingsList).then((result)=>{
    //   setListGUID(result);
    // }); 

  },[currentCalDate]);

  React.useEffect(()=>{
    setEventSources(prevEventSources => reRenderCalendars(prevEventSources, calsVisibility));
  },[JSON.stringify(calsVisibility)]);

  const onLegendChkChange = (calId: string) =>{
    return(ev: any, checked: boolean) =>{
      if (calId !== 'all'){
        setCalsVisibility(prev => {
          const clonePrev = [...prev];
          const newCalsVis = clonePrev.map(item => {
            if (item.calId === calId) item.calChk = checked;
            return item;
          });
          const renderedCalsLen = newCalsVis.filter((item: any) => item.calRender && item.calId !== 'all').length;
          const chkdCalsLen = newCalsVis.filter((item: any) => item.calChk && item.calRender && item.calId !== 'all').length;
          return newCalsVis.map(item => {
            if (item.calId === 'all') item.calChk = renderedCalsLen === chkdCalsLen;
            return item;
          });
        });
      }else{
        setCalsVisibility(prev => {
          const clonePrev = [...prev];
          return clonePrev.map(item => {
            item.calChk = checked;
            return item;
          });
        });
      }
    };
  };
  const chkHandleChange = (newCalSettings:any)=>{    
    return (ev: any, checked: boolean) => { 

      // console.log("newCalSettings", newCalSettings);
      toggleIsDataLoading();
      updateCalSettings(props.context, calSettingsList, newCalSettings, checked).then(()=>{
        _calendarOps.displayCalendars(props.context, calSettingsList, currentCalDate, userGrps, posGrps, Number(props.spCalPageSize), graphCalParams).then((result:{}[])=>{
          // setEventSources(result);
          setEventSources(reRenderCalendars(result, calsVisibility));
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
      updateCalSettings(props.context, calSettingsList, newCalSettings, newCalSettings.ShowCal, item.key, rotaryCals).then(()=>{
        _calendarOps.displayCalendars(props.context, calSettingsList, currentCalDate, userGrps, posGrps, Number(props.spCalPageSize), graphCalParams).then((result:{}[])=>{
          // setEventSources(result);
          setEventSources(reRenderCalendars(result, calsVisibility));
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
    console.log("ev details arg", arg);
    //console.log(formatEvDetails(arg));
    
    if (addedToMyCal.indexOf(arg.event.id) !== -1){
      console.log('Event is already added to your calendar!');
      setEventDetails({...formatEvDetails(arg), EventAdded: true});
    }else{
      setClkdEvId(arg.event.id);
      setEventDetails(formatEvDetails(arg));
    }

    toggleHideDialog();
  };

  const handleAddtoCal = (eventSubject: string, eventBody: string, eventStart: string, eventEnd: string, eventLoc: string)=>{
    console.log("eventSubject", eventSubject);

    addToMyGraphCal(props.context, eventSubject, eventBody, eventStart, eventEnd, eventLoc).then((result)=>{
      const calsAdded = [...addedToMyCal];
      calsAdded.push(clkdEvId);
      setAddedToMyCal(calsAdded);

      console.log('calendar updated', result);
    });
  };

  const passCurrentDate = (currDate: string) => {
    console.log("passCurrentCalDate function", currDate);
    setCurrentCalDate(currDate);
  };

  return(
    <div className={styles.mergedCalendar}>

      {isListView &&
        <Label className={styles.wpTitle}>
          {props.listViewTitle}
        </Label>
      }

      {showLegend && legendPos === 'top' &&
        <div className={`${styles.legendTop} ${legendAlign === 'horizontal' ? styles.legendHz : '' }`}>          
          <ILegend
            calSettings={calSettings} 
            onLegendChkChange={onLegendChkChange}
            legendChked = {calsVisibility}
            userGrps = {userGrps}
            posGrps = {posGrps}
            posGrpView = {props.posGrpView}
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
        isListView = {isListView}
        listViewType = {props.listViewType}
        listViewNavBtns = {props.listViewNavBtns}
        listViewMonthTitle = {props.listViewMonthTitle}
        listViewViews = {props.listViewViews}
        listViewHeight = {props.listViewHeight}
        viewDuration = {props.viewDuration}
        calendarView = {props.calendarView}
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

      {showLegend && legendPos === 'bottom' &&
        <div className={legendAlign === 'horizontal' ? styles.legendHz : '' }>
          <ILegend 
            calSettings={calSettings} 
            onLegendChkChange={onLegendChkChange}
            legendChked = {calsVisibility}
            userGrps = {userGrps}
            posGrps = {posGrps}
            posGrpView = {props.posGrpView}
          />
        </div>
      }

      <IDialog 
        hideDialog={hideDialog} 
        toggleHideDialog={toggleHideDialog}
        eventDetails={eventDetails}
        handleAddtoCal = {handleAddtoCal}
        showAddToCal = {props.showAddToCal}
      />
      
      {/* {showErrors && 
        <div>Test errors! Please ignore!</div>
      } */}
      {showErrors && calMsgErrs.length > 0 &&
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
