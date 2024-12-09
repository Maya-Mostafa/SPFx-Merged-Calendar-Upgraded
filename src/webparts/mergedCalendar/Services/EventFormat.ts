// import * as moment from 'moment';
import * as moment from 'moment-timezone'; 

// only for user's view in the event details dialog, cannot be passed to the fullcalendar plugin
export const formateDate = (ipDate:any) :any => {
    //return moment.utc(ipDate).format('YYYY-MM-DD hh:mm A'); 
    return moment.tz(ipDate, "America/Toronto").format('YYYY-MM-DD hh:mm A');
    // return moment.tz(ipDate, "America/Toronto").format();
};
// only for user's view in the event details dialog
export const formateTime = (ipDate:any) :any => {
    return moment.tz(ipDate, "America/Toronto").format('YYYY-MM-DD hh:mm A');
};

// for fixing the alldayEvent issue in fullcalendar
export const formatStartDate = (ipDate:any) : any => {
    let startDateMod = new Date(ipDate);
    startDateMod.setTime(startDateMod.getTime());
    
    return moment.utc(startDateMod).format('YYYY-MM-DD') + "T" + moment.utc(startDateMod).format("hh:mm") + ":00Z";
    //return moment.tz(startDateMod, "America/Toronto").format('YYYY-MM-DD') + "T" + moment.tz(startDateMod, "America/Toronto").format("hh:mm") + ":00Z";
};
// for fixing the alldayEvent issue in fullcalendar
export const formatEndDate = (ipDate:any) :any => {
    let endDateMod = new Date(ipDate);
    endDateMod.setTime(endDateMod.getTime());

    let nextDay = moment.utc(endDateMod).add(1, 'days');
    return moment.utc(nextDay).format('YYYY-MM-DD') + "T" + moment.utc(nextDay).format("hh:mm") + ":00Z";
    //return moment.tz(nextDay, "America/Toronto").format('YYYY-MM-DD') + "T" + moment.tz(nextDay, "America/Toronto").format("hh:mm") + ":00Z";
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
        // Start: event.startStr ? formateDate(event.startStr) : "",
        // End: event.endStr ? formateDate(event.endStr) : "",
        Start: event._def.extendedProps._startTime ? event._def.extendedProps._startTime : event.start,
        End: event._def.extendedProps._endTime ? event._def.extendedProps._endTime : event.end,
        RecurrenceEnd: event._def.extendedProps._recurrentEndTime,
        Location: event._def.extendedProps._location,
        Body: event._def.extendedProps._body ? event._def.extendedProps._body : null,
        AllDay: event.allDay,
        Recurr: event._def.extendedProps.recurr,
        RecurrData: event._def.extendedProps.recurrData,
        RecurringDef: event._def.extendedProps.recurringDef,
        Category: event._def.extendedProps.category ? (event._def.extendedProps.category.results ? event._def.extendedProps.category.results.join(', ') : null) : null,
        Calendar: event._def.extendedProps.calendar,
        Color: event._def.extendedProps.calendarColor,
        ForeColor: event._def.extendedProps.calendarFontColor,
        EventDayStart: event.start,
        EventDayEnd: event.end,
        EventAdded: false
    };

    return evDetails;
};

export const getDatesRange = (numMonthsStart: number, numMonthsEnd: number) =>{
    const rangeStart = moment().subtract(numMonthsStart, 'months').toISOString();
    const rangeEnd = moment().add(numMonthsEnd, 'months').toISOString();
    
    return {rangeStart: rangeStart, rangeEnd: rangeEnd};
};

export const getDatesWindow = (currentDate: string) => {
    const currentDateVal = new Date (currentDate);
    let dateRangeStart = new Date (currentDate), dateRangeEnd = new Date (currentDate);
    if (currentDateVal.getMonth() === 0){
        dateRangeStart.setMonth(10);
        dateRangeStart.setFullYear(currentDateVal.getFullYear()-1);
    }else{
        dateRangeStart.setMonth(currentDateVal.getMonth()-3);
    }
    if(currentDateVal.getMonth() === 11){
        dateRangeEnd.setMonth(3);
        dateRangeEnd.setFullYear(currentDateVal.getFullYear()+1);
    }else{
        dateRangeEnd.setMonth(currentDateVal.getMonth()+3);
    }

    // console.log("resolveCalUrl current currentDate", currentDate);
    // console.log("currentDate", new Date(currentDate));
    // console.log("dateRangeStart", dateRangeStart);
    // console.log("dateRangeEnd", dateRangeEnd);

    return {dateRangeStart, dateRangeEnd};
};


