import * as moment from 'moment-timezone';

const getElemAttrs = (el:any) :string[] => {
    let attributesArr :string[] = [];
    for (let i = 0; i < el.attributes.length; i++){
        attributesArr.push(el.attributes[i].nodeName);
    }
    return attributesArr;
};

const getWeekDay = (tagAttrs:string[]) : number => {
    let weekDay:number;
    for(let i=0; i<tagAttrs.length; i++){
        switch (tagAttrs[i]) {
            case ('mo'):
            case ("monday"):
                weekDay = 0;
                break;
            case ('tu'):
            case ("tuesday"):
                weekDay = 1;
                break;
            case ('we'):
            case ("wednesday"):
                weekDay = 2;
                break;
            case ('th'):
            case ("thursday"):
                weekDay = 3;
                break;
            case ('fr'):
            case ("friday"):
                weekDay = 4;
                break;
            case ('sa'):
            case ("saturday"):
                weekDay = 5;
                break;
            case ('su'):
            case ("sunday"):
                weekDay = 6;
                break;
        }
    }
    return weekDay;
};

const getWeekDays = (tagAttrs:string[]) : number[] => {
    let weekDay:number = -1,
        weekDays: number[] = [];
    for(let i=0; i<tagAttrs.length; i++){
        switch (tagAttrs[i]) {
            case ('mo'):
            case ("monday"):
                weekDay = 0;
                break;
            case ('tu'):
            case ("tuesday"):
                weekDay = 1;
                break;
            case ('we'):
            case ("wednesday"):
                weekDay = 2;
                break;
            case ('th'):
            case ("thursday"):
                weekDay = 3;
                break;
            case ('fr'):
            case ("friday"):
                weekDay = 4;
                break;
            case ('sa'):
            case ("saturday"):
                weekDay = 5;
                break;
            case ('su'):
            case ("sunday"):
                weekDay = 6;
                break;
            case ('weekday'):
                weekDays = [0, 1, 2, 3, 4];
                break;
        }
        if(weekDay != -1)
            weekDays.push(weekDay);
    }
    return weekDays;
};

const getDayOrder = (weekDayOfMonth:any):number => {
    let dayOrder:number;
    switch (weekDayOfMonth) {
        case ("first"):
            dayOrder = 1;
            break;
        case ("second"):
            dayOrder = 2;
            break;
        case ("third"):
            dayOrder = 3;
            break;
        case ("fourth"):
            dayOrder = 4;
            break;
        case ("last"):
            dayOrder = -1;
            break;
    }
    return dayOrder;
};

const getFirstDayOfWeek = (firstDayOfWeek: string) =>{
    let firstDayOfWeekIndex = 6;
    switch(firstDayOfWeek){
        case "mo" :
        case "monday":
            firstDayOfWeekIndex = 0;
            break;
        case "tu":
        case "tuesday":
            firstDayOfWeekIndex = 1;
            break;
        case "we":
        case "wednesday":
            firstDayOfWeekIndex = 2;
            break;
        case "th":
        case "thursday":
            firstDayOfWeekIndex = 3;
            break;
        case "fr":
        case "friday":
            firstDayOfWeekIndex = 4;
            break;
        case "sa":
        case "saturday":
            firstDayOfWeekIndex = 5;
            break;
        case "su":
        case "sunday":
            firstDayOfWeekIndex = 6;
            break;
    }
    return firstDayOfWeekIndex;
};

const mapGraphToRRuleFreq = (freq: string) =>{
    let freqMapped = freq;
    switch(freq){
        case "relativeMonthly":
        case "absoluteMonthly":
            freqMapped = "monthly";
            break;
        case "relativeYearly":
        case "absoluteYearly":
            freqMapped = "yearly";
            break;
    }
    return freqMapped;
}

export const parseRecurrentEvent = (recurrXML:string, startDate:string, endDate:string) : {} =>{
    
    // for daylight savings
    const modStartDate = moment.tz(startDate, "America/Toronto").format('YYYY-MM-DD') + "T" + moment.tz(startDate, "America/Toronto").format("hh:mm:ss") ;
    const modEndDate = moment.tz(endDate, "America/Toronto").format('YYYY-MM-DD') + "T" + moment.tz(endDate, "America/Toronto").format("hh:mm:ss") ;

    let rruleObj
        : {  wkst: number, dtstart: string, until: string, count: number, interval: number, freq: string, bymonth: number[], bymonthday: number[], byweekday: {}[], bysetpos: number[] }
        = {  wkst: 6, dtstart: modStartDate, until: modEndDate, count: null, interval: 1, freq: null, bymonth: null, bymonthday: null, byweekday: null, bysetpos: null };

// : { tzid: string, wkst: number, dtstart: string, until: string, count: number, interval: number, freq: string, bymonth: number[], bymonthday: number[], byweekday: {}[], bysetpos: number[] }
// = { tzid: "America/Toronto", wkst: 6, dtstart: startDate, until: endDate, count: null, interval: 1, freq: null, bymonth: null, bymonthday: null, byweekday: null, bysetpos: null };


    if (recurrXML.indexOf("<recurrence>") != -1) {
        let $recurrTag : HTMLElement = document.createElement("div");
        $recurrTag.innerHTML = recurrXML;
        
        //console.log($recurrTag);
        const firstDayOfWeek = $recurrTag.getElementsByTagName('firstDayOfWeek')[0].textContent;
        rruleObj.wkst = getFirstDayOfWeek(firstDayOfWeek);

        switch (true) {
            //yearly
            case ($recurrTag.getElementsByTagName('yearly').length != 0):                
                let $yearlyTag = $recurrTag.getElementsByTagName('yearly')[0];
                rruleObj.freq = "yearly";        
                rruleObj.interval = parseInt($yearlyTag.getAttribute('yearfrequency'));
                rruleObj.bymonth = [parseInt($yearlyTag.getAttribute('month'))];
                rruleObj.bymonthday = [parseInt($yearlyTag.getAttribute('day'))];
                break;

            //yearly by day
            case ($recurrTag.getElementsByTagName('yearlybyday').length != 0):
                let $yearlybydayTag = $recurrTag.getElementsByTagName('yearlybyday')[0];
                rruleObj.freq = "yearly";
                rruleObj.interval = parseInt($yearlybydayTag.getAttribute('yearfrequency'));
                rruleObj.bymonth = [parseInt($yearlybydayTag.getAttribute('month'))];

                //attribute mo=TRUE or su=TRUE etc.
                if ($yearlybydayTag.getAttribute('mo') || 
                    $yearlybydayTag.getAttribute('tu') ||
                    $yearlybydayTag.getAttribute('we') ||
                    $yearlybydayTag.getAttribute('th') ||
                    $yearlybydayTag.getAttribute('fr')){
                        rruleObj.byweekday = [{
                            weekday: getWeekDay(getElemAttrs($yearlybydayTag)), 
                            n: getDayOrder($yearlybydayTag.getAttribute('weekdayofmonth'))
                        }]; 
                    }
                
                //attribute day=TRUE
                if($yearlybydayTag.getAttribute('day')){
                    rruleObj.bymonthday = [getDayOrder($yearlybydayTag.getAttribute('weekdayofmonth'))];
                }

                //attribute weekday=TRUE
                if($yearlybydayTag.getAttribute('weekday')){
                    rruleObj.bysetpos = [getDayOrder($yearlybydayTag.getAttribute('weekdayofmonth'))];
                    rruleObj.byweekday = [0,1,2,3,4]; 
                }

                //attribute weekend_day=TRUE
                if($yearlybydayTag.getAttribute('weekend_day')){
                    rruleObj.bysetpos = [getDayOrder($yearlybydayTag.getAttribute('weekdayofmonth'))];
                    rruleObj.byweekday = [5,6]; 
                }
                break;

            //monthly
            case ($recurrTag.getElementsByTagName('monthly').length != 0):
                let $monthlyTag = $recurrTag.getElementsByTagName('monthly')[0];
                rruleObj.freq = "monthly";
                rruleObj.interval = parseInt($monthlyTag.getAttribute('monthfrequency'));
                rruleObj.bymonthday = $monthlyTag.getAttribute('day') ? [parseInt($monthlyTag.getAttribute('day'))]: null;
                break;

            //monthly by day
            case ($recurrTag.getElementsByTagName('monthlybyday').length != 0):
                let $monthlybydayTag = $recurrTag.getElementsByTagName('monthlybyday')[0];
                rruleObj.freq = "monthly";
                rruleObj.interval = parseInt($monthlybydayTag.getAttribute('monthfrequency'));
                
                //attribute mo=TRUE or su=TRUE etc.
                if ($monthlybydayTag.getAttribute('mo') || 
                    $monthlybydayTag.getAttribute('tu') ||
                    $monthlybydayTag.getAttribute('we') ||
                    $monthlybydayTag.getAttribute('th') ||
                    $monthlybydayTag.getAttribute('fr')){
                        rruleObj.byweekday = [{
                            weekday: getWeekDay(getElemAttrs($monthlybydayTag)), 
                            n: getDayOrder($monthlybydayTag.getAttribute('weekdayofmonth'))
                        }]; 
                    }

                //attribute day=TRUE
                if($monthlybydayTag.getAttribute('day'))
                    rruleObj.bymonthday = [getDayOrder($monthlybydayTag.getAttribute('weekdayofmonth'))];
                
                //attribute weekday=TRUE
                if($monthlybydayTag.getAttribute('weekday')){
                    rruleObj.bysetpos = [getDayOrder($monthlybydayTag.getAttribute('weekdayofmonth'))];
                    rruleObj.byweekday = [0,1,2,3,4]; 
                }

                //attribute weekend_day=TRUE
                if($monthlybydayTag.getAttribute('weekend_day')){
                    rruleObj.bysetpos = [getDayOrder($monthlybydayTag.getAttribute('weekdayofmonth'))];
                    rruleObj.byweekday = [5,6]; 
                }
                break;

            //weekly
            case ($recurrTag.getElementsByTagName('weekly').length != 0):
                let $weeklyTag = $recurrTag.getElementsByTagName('weekly')[0];
                rruleObj.freq = "weekly";
                rruleObj.interval = parseInt($weeklyTag.getAttribute('weekfrequency'));
                rruleObj.byweekday = getWeekDays(getElemAttrs($weeklyTag));
                break;

            //daily
            case ($recurrTag.getElementsByTagName('daily').length != 0):
                let $dailyTag = $recurrTag.getElementsByTagName('daily')[0];
                rruleObj.freq = "daily";
                rruleObj.interval = $dailyTag.getAttribute('dayfrequency') ? parseInt($dailyTag.getAttribute('dayfrequency')): 1;
                rruleObj.byweekday = getWeekDays(getElemAttrs($dailyTag));
                break;
        }

        if ($recurrTag.getElementsByTagName('repeatInstances').length != 0)
            rruleObj.count = parseInt($recurrTag.getElementsByTagName('repeatInstances')[0].innerHTML);
        
        //console.log("rruleObj", rruleObj);

        return rruleObj;
        //return { dtstart: startDate, until: endDate, freq: "daily", interval: 1 }

    } else return { dtstart: startDate, until: endDate, freq: "daily", interval: 1 };
};

export const parseGraphRecurrentEv = (graphRecurrenceObj: any, startDateTime: string, endDateTime: string, eventTitle: string) => {
    /*
    "recurrence": {
        "pattern": {
            "type": "weekly",
            "interval": 1,
            "month": 0,
            "dayOfMonth": 0,
            "daysOfWeek": [
                "tuesday",
                "thursday"
            ],
            "firstDayOfWeek": "sunday",
            "index": "first"
        },
        "range": {
            "type": "endDate",
            "startDate": "2024-10-29",
            "endDate": "2025-01-21",
            "recurrenceTimeZone": "Eastern Standard Time",
            "numberOfOccurrences": 0
        }
    },
    */

    console.log("graphRecurrenceObj", graphRecurrenceObj);

    const startDate = graphRecurrenceObj.range.startDate;
    const endDate = graphRecurrenceObj.range.endDate === "0001-01-01" ? "2174-01-01" : graphRecurrenceObj.range.endDate;

    // for daylight savings
    const modStartDate = moment.tz(startDate, "America/Toronto").format('YYYY-MM-DD') + "T" + moment.tz(startDateTime, "America/Toronto").format("hh:mm:ss") ;
    const modEndDate = moment.tz(endDate, "America/Toronto").format('YYYY-MM-DD') + "T" + moment.tz(endDateTime, "America/Toronto").format("hh:mm:ss") ;


    let rruleObj
        : {  wkst: number, dtstart: string, until: string, count: number, interval: number, freq: string, bymonth: number[], bymonthday: number[], byweekday: {}[], bysetpos: number[] }
        = {  wkst: 6, dtstart: modStartDate, until: modEndDate, count: null, interval: 1, freq: '', bymonth: null, bymonthday: null, byweekday: null, bysetpos: null };


    rruleObj.wkst = getFirstDayOfWeek(graphRecurrenceObj.pattern.firstDayOfWeek);
    rruleObj.interval = graphRecurrenceObj.pattern.interval;
    rruleObj.freq = mapGraphToRRuleFreq(graphRecurrenceObj.pattern.type);
    
    if (graphRecurrenceObj.pattern.type === 'relativeMonthly'){
        rruleObj.byweekday = [{
            weekday: getWeekDay(graphRecurrenceObj.pattern.daysOfWeek), 
            n: getDayOrder(graphRecurrenceObj.pattern.index)
        }]; 
    }
    if (graphRecurrenceObj.pattern.type === "absoluteYearly"){        
        rruleObj.bymonthday = [parseInt(graphRecurrenceObj.pattern.dayOfMonth)];
        rruleObj.bymonth = [parseInt(graphRecurrenceObj.pattern.month)];
    }
    if (graphRecurrenceObj.pattern.type === "weekly"){        
        rruleObj.byweekday = getWeekDays(graphRecurrenceObj.pattern.daysOfWeek);
    }


    // else if (graphRecurrenceObj.pattern.type === 'monthly'){ // not tested
    //     rruleObj.bymonthday = graphRecurrenceObj.pattern.dayOfMonth;
    // }
    // else if (graphRecurrenceObj.pattern.type === 'relativeYearly'){ // not tested
    //     rruleObj.byweekday = [{
    //         weekday: getWeekDay(graphRecurrenceObj.pattern.daysOfWeek), 
    //         n: getDayOrder(graphRecurrenceObj.pattern.index)
    //     }]; 
    //     rruleObj.bymonthday = [getDayOrder(graphRecurrenceObj.pattern.dayOfMonth)];
    //     rruleObj.bymonth = graphRecurrenceObj.pattern.month;
    // }
    // else if(graphRecurrenceObj.pattern.type === 'yearly'){ // not tested
    //     rruleObj.bymonthday = [getDayOrder(graphRecurrenceObj.pattern.dayOfMonth)];
    //     rruleObj.bymonth = graphRecurrenceObj.pattern.month;        
    // }
    // else{
    //     rruleObj.byweekday = getWeekDays(graphRecurrenceObj.pattern.daysOfWeek);
    // }
    
    // rruleObj.bymonth = graphRecurrenceObj.pattern.month;
    // rruleObj.bymonthday = graphRecurrenceObj.pattern.dayOfMonth;
    // rruleObj.count = graphRecurrenceObj.range.numberOfOccurrences;
    
    return rruleObj;
}