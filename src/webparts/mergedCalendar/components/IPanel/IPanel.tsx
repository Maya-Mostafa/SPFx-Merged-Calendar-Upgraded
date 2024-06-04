import * as React from 'react';
import { IPanelProps } from './IPanelProps';
import ICalSettings from '../ICalSettings/ICalSettings';
import IPreloader from '../IPreloader/IPreloader';
import {Panel} from 'office-ui-fabric-react';

export default function IPanel (props:IPanelProps) {

    return(
        <Panel
            isOpen={props.isOpen}
            onDismiss={props.dismissPanel}
            headerText="Calendar Settings"
            closeButtonAriaLabel="Close"
            isFooterAtBottom={true}>
            
            <ICalSettings 
                calSettings={props.calSettings}
                dpdOptions={props.dpdOptions}
                onChkChange={props.onChkChange}
                onDpdChange={props.onDpdChange}
                showWeekends={props.showWeekends}
                onChkViewChange={props.onChkViewChange}
            />
            
            <IPreloader isDataLoading={props.isDataLoading} />

        </Panel>
    );
}