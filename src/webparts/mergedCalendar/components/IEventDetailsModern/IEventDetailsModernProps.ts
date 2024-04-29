export interface IEventDetailsModernProps{
    Title: string;
    Start?: string;
    End?: string;
    AllDay?: string;
    Location?: string;
    Body?: any;
    Recurrence?: string;
    handleAddtoCal: any;
    Category: string;
    CalendarName: string;
    CalendarColor: string;
    EventCalDate: string;
    EventCalEndDate: string;
    CalendarFontColor: string;
}