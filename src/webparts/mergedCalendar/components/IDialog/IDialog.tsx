import * as React from 'react';
import styles from '../MergedCalendar.module.scss';
import '../MergedCalendar.scss';
import {Dialog, DialogType, DialogFooter, DefaultButton} from '@fluentui/react';

import {IDialogProps} from './IDialogProps';
// import IEventDetails from  '../IEventDetails/IEventDetails';
import IEventDetailsModern from '../IEventDetailsModern/IEventDetailsModern';

export default function IDialog(props:IDialogProps){

    //const dlgTitleMkp = <span><span className={styles.evTitleDlg} style={{backgroundColor: props.eventDetails.Color }}></span>{props.eventDetails.Calendar}</span> ;

    const modelProps = {
        isBlocking: false,
        className: 'eventDetailsModernModel'
        //styles: { main: { minWidth: '30%' } },
      };
      const dialogContentProps = {
        type: DialogType.close,
        //title: dlgTitleMkp,
        subText: '',
        
      };

      //console.log("props.eventDetails", props.eventDetails);
  
      return (
        <>
          <Dialog
            hidden={props.hideDialog}
            onDismiss={props.toggleHideDialog}
            dialogContentProps={dialogContentProps}
            modalProps={modelProps}
            minWidth="35%" 
            >

            <IEventDetailsModern 
                Title ={props.eventDetails.Title} 
                Start ={props.eventDetails.Start}
                End = {props.eventDetails.End}
                AllDay = {props.eventDetails.AllDay}
                Body = {props.eventDetails.Body}
                Location = {props.eventDetails.Location}       
                handleAddtoCal = {props.handleAddtoCal}         
                Category = {props.eventDetails.Category}
                CalendarName = {props.eventDetails.Calendar}
                CalendarColor= {props.eventDetails.Color}
                CalendarFontColor = {props.eventDetails.ForeColor}
                EventCalDate = {props.eventDetails.EventDayStart}
                EventCalEndDate = {props.eventDetails.EventDayEnd}
                EventAdded = {props.eventDetails.EventAdded}
            />
            <DialogFooter>
              <DefaultButton onClick={props.toggleHideDialog} text="Close" />
            </DialogFooter>
          </Dialog>
        </>
      );
}