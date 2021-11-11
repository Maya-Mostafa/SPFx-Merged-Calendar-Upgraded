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
                weekDay = 0;
                break;
            case ('tu'):
                weekDay = 1;
                break;
            case ('we'):
                weekDay = 2;
                break;
            case ('th'):
                weekDay = 3;
                break;
            case ('fr'):
                weekDay = 4;
                break;
            case ('sa'):
                weekDay = 5;
                break;
            case ('su'):
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
                weekDay = 0;
                break;
            case ('tu'):
                weekDay = 1;
                break;
            case ('we'):
                weekDay = 2;
                break;
            case ('th'):
                weekDay = 3;
                break;
            case ('fr'):
                weekDay = 4;
                break;
            case ('sa'):
                weekDay = 5;
                break;
            case ('su'):
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


export const parseRecurrentEvent = (recurrXML:string, startDate:string, endDate:string) : {} =>{
    let rruleObj
        : { tzid: string, wkst: number, dtstart: string, until: string, count: number, interval: number, freq: string, bymonth: number[], bymonthday: number[], byweekday: {}[], bysetpos: number[] }
        = { tzid: "America/Toronto", wkst: 6, dtstart: startDate, until: endDate, count: null, interval: 1, freq: null, bymonth: null, bymonthday: null, byweekday: null, bysetpos: null };

    if (recurrXML.indexOf("<recurrence>") != -1) {
        let $recurrTag : HTMLElement = document.createElement("div");
        $recurrTag.innerHTML = recurrXML;

        //console.log($recurrTag)

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