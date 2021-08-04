export interface IEventDetailsProps{
    Title: string;
    Start?: string;
    End?: string;
    AllDay?: string;
    Location?: string;
    Body?: any;
    Recurrence?: string;
    handleAddtoCal: any;
}