import { WebPartContext } from "@microsoft/sp-webpart-base";

import {getCalsData} from '../Services/CalendarRequests';
import {getCalSettings} from '../Services/CalendarSettingsOps';

export class CalendarOperations{
    

    public displayCalendars(context: WebPartContext , calSettingsListName:string, userGrps:any, spCalPageSize?: number, graphCalParams?: {rangeStart: number, rangeEnd: number, pageSize: number}): Promise <{}[]>{
        
        console.log("Display Calendar Function");

        let eventSources : {}[] = [], 
            eventSrc  : {} = {};

        // `async` is needed since we're using `await`
        return getCalSettings(context, calSettingsListName).then(async (settings:any) => {
            
            const dataFetches = settings.map(setting => {
                // This `return` is needed otherwise `undefined` is returned in this `map()` call.
                if(setting.ShowCal){ //&& setting.CalType !== 'External'
                    return getCalsData(context, setting, userGrps, spCalPageSize, graphCalParams).then((events: any) => {
                        eventSrc = {
                            events: events,
                            color: setting.BgColorHex,
                            textColor: setting.FgColorHex,
                            calId: setting.Id
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