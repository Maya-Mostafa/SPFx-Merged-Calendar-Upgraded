// import * as moment from 'moment';
import * as moment from 'moment-timezone'; 

export const formateDate = (ipDate:any) :any => {
    //return moment.utc(ipDate).format('YYYY-MM-DD hh:mm A'); 
    //return ipDate; 
    return moment.tz(ipDate, "America/Toronto").format('YYYY-MM-DD hh:mm A');
};

export const formateTime = (ipDate:any) :any => {
    return moment.tz(ipDate, "America/Toronto").format('hh:mm A');
};

export const formatStartDate = (ipDate:any) : any => {
    let startDateMod = new Date(ipDate);
    startDateMod.setTime(startDateMod.getTime());
    
    //return moment.utc(startDateMod).format('YYYY-MM-DD') + "T" + moment.utc(startDateMod).format("hh:mm") + ":00Z";
    return moment.tz(startDateMod, "America/Toronto").format('YYYY-MM-DD') + "T" + moment.tz(startDateMod, "America/Toronto").format("hh:mm") + ":00Z";
};

export const formatEndDate = (ipDate:any) :any => {
    let endDateMod = new Date(ipDate);
    endDateMod.setTime(endDateMod.getTime());

    let nextDay = moment.tz(endDateMod, "America/Toronto").add(1, 'days');
    //return moment.utc(nextDay).format('YYYY-MM-DD') + "T" + moment.utc(nextDay).format("hh:mm") + ":00Z";
    return moment.tz(nextDay, "America/Toronto").format('YYYY-MM-DD') + "T" + moment.tz(nextDay, "America/Toronto").format("hh:mm") + ":00Z";
};

export const formatStrHtml = (str: string) : any => {
    let parser = new DOMParser();
    let htmlEl = parser.parseFromString(str, 'text/html');
    //console.log(htmlEl.body);
    return htmlEl.body;
};

export const formatEvDetails = (ev:any) : {} =>{
    let event = ev.event,
        evDetails : {} = {};

    evDetails = {
        Title: event.title,
        Start: event.startStr ? formateDate(event.startStr) : "",
        End: event.endStr ? formateDate(event.endStr) : "",
        Location: event._def.extendedProps._location,
        Body: event._def.extendedProps._body ? event._def.extendedProps._body : null,
        AllDay: event.allDay,
        Recurr: event._def.extendedProps.recurr,
        RecurrData: event._def.extendedProps.recurrData,
        RecurringDef: event._def.extendedProps.recurringDef
    };

    return evDetails;
};

export const getDatesRange = (numMonthsStart: number, numMonthsEnd: number) =>{
    const rangeStart = moment().subtract(numMonthsStart, 'months').toISOString();
    const rangeEnd = moment().add(numMonthsEnd, 'months').toISOString();
    
    return {rangeStart: rangeStart, rangeEnd: rangeEnd};
};


