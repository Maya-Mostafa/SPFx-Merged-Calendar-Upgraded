import { WebPartContext } from "@microsoft/sp-webpart-base";
import {HttpClientResponse, HttpClient, IHttpClientOptions, MSGraphClient, SPHttpClient} from "@microsoft/sp-http";

import {formatStartDate, formatEndDate, getDatesRange, formateTime} from '../Services/EventFormat';
import {parseRecurrentEvent} from '../Services/RecurrentEventOps';

import axios, { AxiosRequestConfig } from 'axios';

export const calsErrs : any = [];

export const getPosGrpMapping = (posGrpName: string) => {
    const posGrpsMapping = [];
    posGrpsMapping['SOE'] = [11, 84];
    posGrpsMapping['ASG - Administrative Staff Group'] = [12, 16, 17, 18, 19, 25, 26, 60, 61, 62, 64, 64, 70, 90, 98];
    posGrpsMapping['CUPE 2544 - Custodial, Maintenance and Food Services'] = [50, 51, 52, 55, 56, 59];
    posGrpsMapping['Teachers'] = [30, 31, 32, 33, 34, 35, 20, 21, 22, 23, 24, 91, 92, 96];
    posGrpsMapping['Elementary Teachers'] = [30, 31, 32, 33, 34, 35];
    posGrpsMapping['Secondary Teachers'] = [20, 21, 22, 23, 24, 91, 92, 96];
    posGrpsMapping['CUPE 1628 - Secretarial, Clerical and Library Technicians'] = [40, 41, 42, 47, 48, 49, 86, 87, 88];
    posGrpsMapping['CUPE 1628 - Secretarial'] = [40, 41, 42, 47, 48, 49, 86, 87, 88];
    posGrpsMapping['Clerical and Library Technicians'] = [40, 41, 42, 47, 48, 49, 86, 87, 88];
    posGrpsMapping['School Admin(P-VPs)'] = [28, 29, 38, 39, 82, 83, 90, 98];
    posGrpsMapping['ECE - Educational Credential Assessment'] = [93, 94, 95];
    posGrpsMapping['OPSEU-2100 Educational Assistants/Designated Early Childhood Educators'] = [13, 14, 15];
    posGrpsMapping['OPSEU-2100 Educational Assistants'] = [13, 14, 15];
    posGrpsMapping['Designated Early Childhood Educators'] = [13, 14, 15];
    posGrpsMapping['OPSEU'] = [65, 66, 67, 69, 74, 77, 78, 79, 80, 81];
    posGrpsMapping['Casual'] = [71, 72, 73, 75, 76, 89];

    return posGrpsMapping[posGrpName];
};

export const getAllPosGrps = async (context:WebPartContext) => {
    const posGrpsMapping = [];
    const data = await context.spHttpClient.get("https://pdsb1.sharepoint.com/sites/contentTypeHub/_api/web/lists/getByTitle('PLEmpGrps')/items", SPHttpClient.configurations.v1);
    if (data.ok){
        const results = await data.json();
        const posGrps = results.value;
        for (let posGrp of posGrps){
            posGrpsMapping[posGrp.Title] = posGrp.Numbers.split(';').map(Number);
        }
    }
    return posGrpsMapping;
};

export const getUserGrp = async (context: WebPartContext) => {
    const userEmail = context.pageContext.user.email;
    const empListRespUrl = `https://pdsb1.sharepoint.com/sites/contentTypeHub/_api/web/lists/getByTitle('Employees')/items?$filter=MMHubBoardEmail eq '${userEmail}'&$select=MMHubEmployeeGroup'`;
    const empListResp = await context.spHttpClient.get(empListRespUrl, SPHttpClient.configurations.v1);

    if (empListResp.ok){
        const results = await empListResp.json();
        if (results.value[0] && results.value[0].MMHubEmployeeGroup)
            return results.value[0].MMHubEmployeeGroup.split(';').filter(item => Number(item));
        return [];
    }
};

const resolveCalUrl = (context: WebPartContext, calType:string, calUrl:string, calName:string, currentDate: string, spCalPageSize?: number) : string => {
    
    let resolvedCalUrl:string,
        azurePeelSchoolsUrl :string = "https://pdsb1.azure-api.net/peelschoolstemp",
        restApiUrl :string = "/_api/web/lists/getByTitle('"+calName+"')/items",
        restApiUrlExt :string = "/_api/web/lists/getByTitle('School - Calendar')/items",
        //restApiParams :string = "?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData&$filter=EventDate ge datetime'2019-08-01T00%3a00%3a00'";
        //restApiParams :string = "?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData&$orderby=EventDate desc&$top=300";
        
        //restApiParams :string = `?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData&$top=${spCalParams.pageSize}&$filter=EventDate ge '${getDatesRange(spCalParams.rangeStart, spCalParams.rangeEnd).rangeStart}' and EventDate le '${getDatesRange(spCalParams.rangeStart, spCalParams.rangeEnd).rangeEnd}' or EndDate ge '${getDatesRange(spCalParams.rangeStart, spCalParams.rangeEnd).rangeStart}'`;
        restApiParams :string = `?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData,Category&$top=${spCalPageSize}&$orderby=EndDate desc`,
        restApiParamsExt :string = `?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData&$top=${spCalPageSize}&$orderby=EndDate desc`;
    //$filter=EventDate ge datetime'2019-08-01T00%3a00%3a00'

    /*
    get the prev and next month dates from the current date --> dateRange
    if not recurrent
        display event in the current dateRange 
            -> startDate in dateRange or endDate in dateRange
            -> i.e. date >= startDateRange || date <= endDateRange
    if recurrent display all
     */

    const currentDateVal = new Date (currentDate);
    let dateRangeStart = new Date (currentDate), dateRangeEnd = new Date (currentDate);
    if (currentDateVal.getMonth() === 0){
        dateRangeStart.setMonth(11);
        dateRangeStart.setFullYear(currentDateVal.getFullYear()-1);
    }else{
        dateRangeStart.setMonth(currentDateVal.getMonth()-3);
    }
    if(currentDateVal.getMonth() === 11){
        dateRangeEnd.setMonth(0);
        dateRangeEnd.setFullYear(currentDateVal.getFullYear()+1);
    }else{
        dateRangeEnd.setMonth(currentDateVal.getMonth()+3);
    }

    console.log("resolveCalUrl current currentDate", currentDate);
    console.log("currentDate", new Date(currentDate));
    console.log("dateRangeStart", dateRangeStart);
    console.log("dateRangeEnd", dateRangeEnd);

    let restApiParamsWRange :string = `?$select=ID,Title,EventDate,EndDate,Location,Description,fAllDayEvent,fRecurrence,RecurrenceData,Category&$top=${spCalPageSize}&$orderby=EndDate desc&$filter=fRecurrence eq 1 or EventDate ge '${dateRangeStart.toISOString()}' and EventDate le '${dateRangeEnd.toISOString()}'`;
    restApiParams = restApiParamsWRange;

    switch (calType){
        case "Internal":
        case "Rotary":
            resolvedCalUrl = calUrl + restApiUrl + restApiParams;
            break;
        case "My School":
            resolvedCalUrl = context.pageContext.web.absoluteUrl + restApiUrl + restApiParams;
            break;
        case "External":
            // resolvedCalUrl = azurePeelSchoolsUrl + calUrl.substring(calUrl.indexOf('.org/') + 12, calUrl.length) + restApiUrlExt + restApiParamsExt;
            resolvedCalUrl = calUrl;
            break;
    }
    return resolvedCalUrl;
};

const getGraphCals = (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string, BgColorHex: string}, currentDate: string, graphCalParams?: {rangeStart: string, rangeEnd: string, pageSize: string}) : Promise <{}[]> => {
    
    let graphUrl :string = calSettings.CalURL.substring(32, calSettings.CalURL.length),
        calEvents : {}[] = [];
    
    const currentDateVal = new Date (currentDate);
    let dateRangeStart = new Date (currentDate), dateRangeEnd = new Date (currentDate);
    if (currentDateVal.getMonth() === 0){
        dateRangeStart.setMonth(11);
        dateRangeStart.setFullYear(currentDateVal.getFullYear()-1);
    }else{
        dateRangeStart.setMonth(currentDateVal.getMonth()-1);
    }
    if(currentDateVal.getMonth() === 11){
        dateRangeEnd.setMonth(0);
        dateRangeEnd.setFullYear(currentDateVal.getFullYear()+1);
    }else{
        dateRangeEnd.setMonth(currentDateVal.getMonth()+2);
    }

    return new Promise <{}[]> (async(resolve, reject)=>{
        context.msGraphClientFactory
            .getClient()
            .then((client :MSGraphClient)=>{
                client
                    // .api(graphUrl)
                    //.api(`${graphUrl}?$filter=start/dateTime ge '${getDatesRange(Number(graphCalParams.rangeStart), Number(graphCalParams.rangeEnd)).rangeStart}' and start/dateTime le '${getDatesRange(Number(graphCalParams.rangeStart), Number(graphCalParams.rangeEnd)).rangeEnd}'&$top=${Number(graphCalParams.pageSize)}`)
                    .api(`${graphUrl}?$filter=start/dateTime ge '${dateRangeStart.toISOString()}' and start/dateTime le '${dateRangeEnd.toISOString()}'&$top=${Number(graphCalParams.pageSize)}`)
                    .header('Prefer','outlook.timezone="Eastern Standard Time"')
                    .get((error, response: any, rawResponse?: any)=>{
                        if(error){
                            calsErrs.push("MS Graph Error - " + calSettings.Title);
                        }
                        if(response){
                            console.log("graph response", response);
                            response.value.map((result:any)=>{
                                calEvents.push({
                                    id: result.id,
                                    title: result.subject,
                                    // start: formatStartDate(result.start.dateTime),
                                    // end: formatStartDate(result.end.dateTime),
                                    start: result.start.dateTime,
                                    end: result.end.dateTime,
                                    _location: result.location.displayName,
                                    _body: result.body.content,
                                    allDay: result.isAllDay,
                                    calendar: calSettings.Title,
                                    calendarColor: calSettings.BgColorHex
                                });
                            });
                        }
                        resolve(calEvents);
                    });
            }, (error)=>{
                calsErrs.push(error);
            });
    });
};

export const addToMyGraphCal = async (context: WebPartContext) =>{
    
    const event = {
        "subject": "Let's add this to my calendar",
        "body": {
            "contentType": "HTML",
            "content": "Adding a dummy event to my graph calendar"
        },
        "start": {
            "dateTime": "2021-02-15T12:00:00",
            "timeZone": "Pacific Standard Time"
        },
        "end": {
            "dateTime": "2021-02-15T14:00:00",
            "timeZone": "Pacific Standard Time"
        },
        "location": {
            "displayName": "Peel CBO"
        },
        "attendees": [{
            "emailAddress": {
                "address": "mai.mostafa@peelsb.com",
                "name": "Mai Mostafa"
            },
            "type": "required"
        }]
    };

    context.msGraphClientFactory
        .getClient()
        .then((client :MSGraphClient)=>{
            client
                .api("/me/events")
                .post(event, (err, res) => {
                    console.log(res);
                });
        });

};

const getDefaultCals1 = (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string}, currentDate: string) : Promise <{}[]> =>{
    
    let calUrl :string = resolveCalUrl(context, calSettings.CalType, calSettings.CalURL, calSettings.CalName, currentDate),
        calEvents : {}[] = [] ;

    const myOptions: IHttpClientOptions = {
        headers : { 
            'Accept': 'application/json;odata=verbose'
        }
    };

    // console.log("calURL", calUrl);

    return new Promise <{}[]> (async(resolve, reject) =>{
        context.httpClient
            .get(calUrl, HttpClient.configurations.v1, myOptions)
            .then((response: HttpClientResponse) =>{
                response.json().then((results:any)=>{
                    results.d.results.map((result:any)=>{
                        calEvents.push({
                            id: result.ID,
                            title: result.Title,
                            start: result.fAllDayEvent ? formatEndDate(result.EventDate) : result.EventDate,
                            end: result.fAllDayEvent ? formatEndDate(result.EndDate) : result.EndDate,
                            allDay: result.fAllDayEvent,
                            _location: result.Location,
                            _body: result.Description,
                            recurr: result.fRecurrence,
                            recurrData: result.RecurrenceData,
                            rrule: result.fRecurrence ? parseRecurrentEvent(result.RecurrenceData, formatStartDate(result.EventDate), formatEndDate(result.EndDate)) : null
                        });
                    });
                    resolve(calEvents);
                });
            }).catch((error:any)=>{
                resolve([]);
                console.log("Calendar URL error!");
            });
    });
    
};

export const getDefaultCals = async (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string, Id: string, View: string, BgColorHex: string}, currentDate: string, userGrps: [], posGrps:any, spCalPageSize?: number) : Promise <{}[]> => {
    
    let calUrl :string = resolveCalUrl(context, calSettings.CalType, calSettings.CalURL, calSettings.CalName, currentDate, spCalPageSize),
        calEvents : {}[] = [] ;

    const myOptions: IHttpClientOptions = {
        headers : { 
            'Accept': 'application/json;odata=verbose',
        }
    };

    try{
        const _data = await context.httpClient.get(calUrl, HttpClient.configurations.v1, myOptions);
        //console.log(calSettings.Title, _data.status);
        if (_data.ok){
            const calResult = await _data.json();
            // console.log("calResult", calResult);
            // console.log("calSettings view", calSettings.View);

            if(calResult){
                if (calSettings.View && calSettings.View.toLocaleLowerCase() !== 'allitems'){ //for calendars with views

                    // console.log("calSettings.View", calSettings.View);
                    // console.log("calSettings.Title", calSettings.Title);
                    // console.log("userGrps passed here", userGrps);

                    let isUserGrpCal = false;
                    if (posGrps[calSettings.View.trim()] == undefined) isUserGrpCal = true;
                    else{
                        for (let userGrp of userGrps){
                            if (posGrps[calSettings.View.trim()] && posGrps[calSettings.View.trim()].indexOf(Number(userGrp)) !== -1){
                                isUserGrpCal = true;
                                break;
                            }
                        }
                    }
                    
                    // console.log("posGrps in calendar requests", posGrps);
                    // console.log("posGrps[calSettings.Title]", posGrps[calSettings.Title]);

                    calResult.d.results.map((result:any)=>{
                        if (result.Category){
                            if (result.Category === calSettings.View || ( result.Category.results && result.Category.results.indexOf(calSettings.View) !== -1)){
                                calEvents.push({
                                    id: result.ID,
                                    title: result.Title,
                                    start: result.fAllDayEvent ? formatStartDate(result.EventDate) : result.EventDate,
                                    end: result.fAllDayEvent ? formatEndDate(result.EndDate) : result.EndDate,
                                    _startTime: formateTime(result.EventDate),
                                    _endTime: formateTime(result.EndDate),
                                    allDay: result.fAllDayEvent,
                                    _location: result.Location,
                                    _body: result.Description,
                                    recurr: result.fRecurrence,
                                    recurrData: result.RecurrenceData,
                                    rrule: result.fRecurrence ? parseRecurrentEvent(result.RecurrenceData, result.fAllDayEvent ? formatStartDate(result.EventDate) : result.EventDate, result.fAllDayEvent ? formatEndDate(result.EndDate) : result.EndDate) : null,
                                    // className: calVisibility.calId ? ( calVisibility.calId == calSettings.Id && !calVisibility.calChk ? 'eventHidden' : '') : ''
                                    //className: 'eventCal' + calSettings.Id,
                                    className: !isUserGrpCal ? 'eventHidden' : '',
                                    category: result.Category,
                                    calendar: calSettings.Title,
                                    calendarColor: calSettings.BgColorHex
                                });
                            }
                        }
                    });
                }
                else{
                    calResult.d.results.map((result:any)=>{
                        calEvents.push({
                            id: result.ID,
                            title: result.Title,
                            start: result.fAllDayEvent ? formatStartDate(result.EventDate) : result.EventDate,
                            end: result.fAllDayEvent ? formatEndDate(result.EndDate) : result.EndDate,
                            _startTime: formateTime(result.EventDate),
                            _endTime: formateTime(result.EndDate),
                            allDay: result.fAllDayEvent,
                            _location: result.Location,
                            _body: result.Description,
                            recurr: result.fRecurrence,
                            recurrData: result.RecurrenceData,
                            rrule: result.fRecurrence ? parseRecurrentEvent(result.RecurrenceData, result.fAllDayEvent ? formatStartDate(result.EventDate) : result.EventDate, result.fAllDayEvent ? formatEndDate(result.EndDate) : result.EndDate) : null,
                            // className: calVisibility.calId ? ( calVisibility.calId == calSettings.Id && !calVisibility.calChk ? 'eventHidden' : '') : ''
                            //className: 'eventCal' + calSettings.Id,
                            // className: 'eventHidden',
                            category: result.Category,
                            calendar: calSettings.Title,
                            calendarColor: calSettings.BgColorHex
                        });
                    });
                }
            }
        }else{
            calsErrs.push(calSettings.Title + ' - ' + _data.statusText);
            return [];
        }
    } catch(error){
        calsErrs.push("External calendars invalid - " + error);
    }

    // console.log("calSettings", calSettings);
    console.log("getDefaultCals calEvents", calEvents);

    return calEvents;
};

export const getExtCals = async (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string, Id: string, View: string, BgColorHex: string}, currentDate: string, spCalPageSize?: number) : Promise <{}[]> => {
    //let restApiParamsNewExt = calSettings.CalURL.charAt(calSettings.CalURL.length-1) === '/' ? 'api/v1/event' : '/api/v1/event';
    // let calUrl :string = calSettings.CalURL + restApiParamsNewExt;
    
    const currentDateVal = new Date (currentDate);
    let dateRangeStart = new Date (currentDate), dateRangeEnd = new Date (currentDate);
    if (currentDateVal.getMonth() === 0){
        dateRangeStart.setMonth(11);
        dateRangeStart.setFullYear(currentDateVal.getFullYear()-1);
    }else{
        dateRangeStart.setMonth(currentDateVal.getMonth()-3);
    }
    if(currentDateVal.getMonth() === 11){
        dateRangeEnd.setMonth(0);
        dateRangeEnd.setFullYear(currentDateVal.getFullYear()+1);
    }else{
        dateRangeEnd.setMonth(currentDateVal.getMonth()+3);
    }

    let calUrl :string = `${calSettings.CalURL}&startdate=${dateRangeStart}&enddate=${dateRangeEnd}`;
    let calEvents : {}[] = [] ;

    try{
        const _data = await context.httpClient.get(calUrl, HttpClient.configurations.v1);
        if (_data.ok){
            const calResult = await _data.json();
            if(calResult){
                console.log("new external cal results", calResult);
                calResult.map((result:any)=>{
                    calEvents.push({
                        id: result.id,
                        title: result.title,
                        start: new Date(result.settings.startdate).toISOString(),
                        end: new Date(result.settings.enddate).toISOString(),
                        _startTime: formateTime(result.settings.startdate),
                        _endTime: formateTime(result.settings.enddate),
                        _body: result.content,
                        calendar: calSettings.Title,
                        calendarColor: calSettings.BgColorHex,

                        allDay: false,
                        _location: null,
                        recurr: false,
                        //recurrData: result.RecurrenceData,
                        //rrule: result.fRecurrence ? parseRecurrentEvent(result.RecurrenceData, result.fAllDayEvent ? formatStartDate(result.EventDate) : result.EventDate, result.fAllDayEvent ? formatEndDate(result.EndDate) : result.EndDate) : null,
                        
                    });
                });
                console.log("formatted new ext calEvents", calEvents);
            }
        }
    } catch(error){
        calsErrs.push("New External calendars invalid - " + error);
    }
    return calEvents;
};

export const getCalsData2 = (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string, Id: string, View:string, BgColorHex: string}, currentDate: string, userGrps: [], posGrps: any, spCalPageSize?: number, graphCalParams?: {rangeStart: string, rangeEnd: string, pageSize: string}) : Promise <{}[]> => {
    if(calSettings.CalType == 'Graph'){
        return getGraphCals(context, calSettings, currentDate, graphCalParams);
    }else if(calSettings.CalType === 'External' && ((calSettings.CalURL.toLowerCase().indexOf('www.peelschools.org') === -1 && calSettings.CalURL.toLowerCase().indexOf('www2.peelschools.org') === -1))){
        return getExtCals(context, calSettings, currentDate, spCalPageSize);
    }
    else{
        return getDefaultCals(context, calSettings, currentDate, userGrps, posGrps, spCalPageSize);
    }
};

export const getCalsData = (context: WebPartContext, calSettings:{CalType:string, Title:string, CalName:string, CalURL:string, Id: string, View:string, BgColorHex: string}, currentDate: string, userGrps: [], posGrps: any, spCalPageSize?: number, graphCalParams?: {rangeStart: string, rangeEnd: string, pageSize: string}) : Promise <{}[]> => {
    if(calSettings.CalType == 'Graph'){
        return getGraphCals(context, calSettings, currentDate, graphCalParams);
    }else if ( calSettings.CalType == 'External'){
        return getExtCals(context, calSettings, currentDate, spCalPageSize);
    }else{
        return getDefaultCals(context, calSettings, currentDate, userGrps, posGrps, spCalPageSize);
    }
};

export const reRenderCalendars = (calEventSources: any, calVisibility: {calId: string, calChk: boolean}) =>{

    const newCalEventSources = calEventSources.map((eventSource: any) => {
        if (eventSource.calId == calVisibility.calId) {
            const updatedEventSource = {...eventSource}; //shallow clone
            updatedEventSource.events = eventSource.events.map((event: any) => {
                event['className'] = !calVisibility.calChk ? 'eventHidden' : '';
                return event;
            });
            return updatedEventSource;
        } else {
            return {...eventSource}; //shallow clone
        }
    });
    
    // const newCalEventSources = [...calEventSources];
    // for (let i = 0; i < newCalEventSources.length; i++){
    //     if (newCalEventSources[i].calId == calVisibility.calId){
    //         let calEvents = [...newCalEventSources[i]['events']];
    //         for (let j = 0; j< calEvents.length; j++){
    //             calEvents[j]['className'] = 'eventHidden';
    //         }
    //     }
    // }
    
    return newCalEventSources;
};

export const getMySchoolCalGUID = async (context: WebPartContext, calSettingsListName: string) =>{
    const calSettingsRestUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${calSettingsListName}')/items?$filter=CalType eq 'My School'&$select=CalName`;
    const calSettingsCall = await context.spHttpClient.get(calSettingsRestUrl, SPHttpClient.configurations.v1).then(response => response.json());
    const calName = calSettingsCall.value[0].CalName;

    const calRestUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('${calName}')?$select=id`;
    const calCall = await context.spHttpClient.get(calRestUrl, SPHttpClient.configurations.v1).then(response => response.json());
    
    return calCall.Id;
};

