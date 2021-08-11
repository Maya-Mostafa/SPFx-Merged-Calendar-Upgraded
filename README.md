# SPFx-Merged-Calendar
A SPFx Merged Calendar React web-part. Aggregates different types of calendars; internal, external, graph, google using Full Calendar plugin.

Started in March 2020 with just plain JS

Plain JS component was done by October 2020

SPFx with React started in November 2020

November 10 - Adding SPFx files

Cloning to other machine

Screenshots
------------
- Merged Calendar
![alt Calendar](https://github.com/Maya-Mostafa/SPFx-Merged-Calendar-Upgraded/blob/main/mergedCal.png) <br/>
- Legend with hide/show option and links to calendars
![alt Legend](https://github.com/Maya-Mostafa/SPFx-Merged-Calendar-Upgraded/blob/main/legend.png) <br/>
- Calendar Settings
![alt Settings](https://github.com/Maya-Mostafa/SPFx-Merged-Calendar-Upgraded/blob/main/settings.png) <br/>
- Event details dialog
![alt Dialog](https://github.com/Maya-Mostafa/SPFx-Merged-Calendar-Upgraded/blob/main/dialog.png) <br/>


Milestones
------------
- FullCalendar Integration with react
- Get calendar information from SP list using Rest API & Display in FullCalendar
- FullCalendar full day event bug resolution
- FullCalendar Recurrent events parsing
- Reading external calendars from Azure API using HttpClient and not SPHttpClient
- Reading Graph calendars and modifying permissions
- Implementing the Settings panel using Fluent UI
- Implementing the Legend component
- Implementing the dialog and event details components

Change Requirements
-------------------
- Popping an error message on an invalid Calendar URL
- Modifying the legend calendar links to read from a stand-alone field "Link" from the list
- Adding a new feature for showing/hiding calendars from the legend

Upgrade
-------
- An updgrade to Node 14 was done, and a new solution has been added.


Terminal Commands
-------------------
npm install rrule

npm install --save @fullcalendar/react @fullcalendar/rrule @fullcalendar/daygrid @fullcalendar/timegrid @fullcalendar/interaction

npm install moment

npm install @fluentui/react

npm install @fluentui/react-hooks

npm install office-ui-fabric-core


gulp package-solution

gulp serve --nobrowser


gulp bundle --ship

gulp package-solution --ship



