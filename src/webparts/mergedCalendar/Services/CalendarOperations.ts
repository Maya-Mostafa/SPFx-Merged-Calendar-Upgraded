import { WebPartContext } from "@microsoft/sp-webpart-base";

import {getCalsData} from '../Services/CalendarRequests';
import {getCalSettings} from '../Services/CalendarSettingsOps';

export class CalendarOperations{
    

    public displayCalendars(context: WebPartContext , calSettingsListName:string): Promise <{}[]>{
        
        console.log("Display Calendar Function");

        let eventSources : {}[] = [], 
            eventSrc  : {} = {};

        // `async` is needed since we're using `await`
        return getCalSettings(context, calSettingsListName).then(async (settings:any) => {
            const dataFetches = settings.map(setting => {
                // This `return` is needed otherwise `undefined` is returned in this `map()` call.
                if(setting.ShowCal){ //&& setting.CalType !== 'External'
                    return getCalsData(context, setting).then((events: any) => {
                        eventSrc = {
                            events: events,
                            color: setting.BgColorHex,
                            textColor: setting.FgColorHex
                        };
                        eventSources.push(eventSrc);
                    });
                }
            });
            await Promise.all(dataFetches);
            
            // The next then takes the eventSources array and it becomes the return value.
            // Its a one-liner so `return` is implicitly known here
        }).then(() => eventSources);
    }

   
    
}