export interface IMergedCalendarState {
  weekendsVisible: boolean;
  eventSources: {}[];
  calSettingsList: string;
  //eventSources: {events:{}[], color: string, textColor: string}[];

}



/*
eventSources: [
  {events : [{ title: 'event 1', date: '2020-11-05' }], color: "red", textColor: "#fff"},
  {events : [{ title: 'event 2', date: '2020-11-18' }], color: "#000", textColor: "#fff"},
  {events : [{ title: 'event 3', date: '2020-11-26' }], color: "#000", textColor: "#fff"},
]

{}[]


eventSources: {events:{title:string, date:string}[], color: string, textColor: string}[];

*/