import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import {SPHttpClient} from '@microsoft/sp-http';

import * as strings from 'MergedCalendarWebPartStrings';
import MergedCalendar from './components/MergedCalendar';
import { IMergedCalendarProps } from './components/IMergedCalendarProps';

export interface IMergedCalendarWebPartProps {
  description: string;  
  showWeekends: boolean;
  calSettingsList: string;
  legendPos: string;
  legendAlign: string;
  spCalParams : {rangeStart: number, rangeEnd: number, pageSize: number};
  graphCalParams : {rangeStart: number, rangeEnd: number, pageSize: number};
}

export default class MergedCalendarWebPart extends BaseClientSideWebPart<IMergedCalendarWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMergedCalendarProps> = React.createElement(
      MergedCalendar,
      {
        description: this.properties.description,
        showWeekends: this.properties.showWeekends,
        context: this.context,
        calSettingsList: this.properties.calSettingsList,
        legendPos : this.properties.legendPos,
        legendAlign: this.properties.legendAlign,
        dpdOptions : [
          { key: 'E1Day', text: '1 Day Cycle' },
          { key: 'E2Day', text: '2 Day Cycle' },
          { key: 'E3Day', text: '3 Day Cycle' },
          { key: 'E4Day', text: '4 Day Cycle' },
          { key: 'E5Day', text: '5 Day Cycle' },
          { key: 'E6Day', text: '6 Day Cycle' },
          { key: 'E7Day', text: '7 Day Cycle' },
          { key: 'E8Day', text: '8 Day Cycle' },
          { key: 'E9Day', text: '9 Day Cycle' },
          { key: 'E10Day', text: '10 Day Cycle' },
        ],
        spCalParams : this.properties.spCalParams,
        graphCalParams: this.properties.graphCalParams
        
      }
    );

    // spCalParams: {rangeStart: this.properties.spCalParams.rangeStart, rangeEnd: this.properties.spCalParams.rangeEnd, pageSize: this.properties.spCalParams.pageSize},
    // graphCalParams: {rangeStart: this.properties.graphCalParams.rangeStart, rangeEnd: this.properties.graphCalParams.rangeEnd, pageSize: this.properties.graphCalParams.pageSize},

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // protected get disableReactivePropertyChanges(): boolean {
  //   return true;
  // }
  // private validateListName(value: string): string {
  //   if (value === null || value.trim().length === 0) {
  //     return 'Provide a list name';
  //   }
  //   if (value.length > 40) {
  //     return 'List name should not be longer than 40 characters';
  //   }
  //   return '';
  // }

  /* Loading Dpd with list names - Start */
  private lists: IPropertyPaneDropdownOption[];
  private async loadLists(): Promise<IPropertyPaneDropdownOption[]> {    
    let listsTitle : any = [];
    try {
      let response = await this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$select=Title&$filter=BaseType eq 0 and BaseTemplate eq 100 and Hidden eq false`, SPHttpClient.configurations.v1);
      if (response.ok) {
        const results = await response.json();
        if(results){
          console.log('results', results);
          results.value.map((result: any)=>{
            listsTitle.push({
              key: result.Title,
              text: result.Title
            });
          });
          return listsTitle;
        }
      }
    } catch (error) {
      return error.message;
    }
  }
  protected onPropertyPaneConfigurationStart(): void {
    if (this.lists) {
      this.render();  
      return;
    }
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'lists');
    this.loadLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => {
        this.lists = listOptions;
        this.context.propertyPane.refresh();
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);        
        this.render();       
      });
  } 
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'listName' && newValue) {
      // push new list value
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
      // refresh the item selector control by repainting the property pane
      this.context.propertyPane.refresh();
      // re-render the web part as clearing the loading indicator removes the web part body
      this.render();      
    }
    else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, oldValue);
    }
  }
  /* Loading Dpd with list names - End */

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneDropdown('calSettingsList', {
                  label : 'Calendar Settings List',
                  options: this.lists,
                  selectedKey : 'CalendarSettings'
                }),
                PropertyPaneCheckbox('showWeekends', {
                  text: "Show Weekends"
                }),
              ]
            },
            {
              groupName: "Legend",
              groupFields: [
                PropertyPaneDropdown('legendPos', {
                  label : 'Position',
                  options: [
                    {key: 'top', text: 'Top'},
                    {key: 'bottom', text: 'Bottom'}
                    // {key: 'both', text: 'Both'}
                  ],
                  selectedKey : 'top'
                }),
                PropertyPaneDropdown('legendAlign', {
                  label : 'Alignment',
                  options: [
                    {key: 'horizontal', text: 'Horizontal'},
                    {key: 'vertical', text: 'Vertical'}
                  ],
                  selectedKey : 'vertical'
                }),
              ]
            },
            {
              groupName: "SharePoint Calendars",
              groupFields: [
                PropertyPaneDropdown('spCalParams.rangeStart', {
                  label : 'Number of months before today',
                  options: [
                    {key: '1', text: '1'},
                    {key: '2', text: '2'},
                    {key: '3', text: '3'},
                    {key: '4', text: '4'},
                    {key: '5', text: '5'},
                    {key: '6', text: '6'},
                  ],
                  selectedKey : '3'
                }),
                PropertyPaneDropdown('spCalParams.rangeEnd', {
                  label : 'Number of months after today',
                  options: [
                    {key: '1', text: '1'},
                    {key: '2', text: '2'},
                    {key: '3', text: '3'},
                    {key: '4', text: '4'},
                    {key: '5', text: '5'},
                    {key: '6', text: '6'},
                    {key: '7', text: '7'},
                    {key: '8', text: '8'},
                    {key: '9', text: '9'},
                    {key: '10', text: '10'},
                    {key: '11', text: '11'},
                    {key: '12', text: '12'},
                  ],
                  selectedKey : '6'
                }),
                PropertyPaneDropdown('spCalParams.pageSize', {
                  label : 'Number of events',
                  options: [
                    {key: '100', text: '100'},
                    {key: '200', text: '200'},
                    {key: '300', text: '300'},
                    {key: '400', text: '400'},
                    {key: '500', text: '500'},
                    {key: '600', text: '600'},
                    {key: '700', text: '700'},
                  ],
                  selectedKey : '100'
                }),
              ]
            },
            {
              groupName: "Graph Calendars",
              groupFields: [
                PropertyPaneDropdown('graphCalParams.rangeStart', {
                  label : 'Number of months before today',
                  options: [
                    {key: '1', text: '1'},
                    {key: '2', text: '2'},
                    {key: '3', text: '3'},
                    {key: '4', text: '4'},
                    {key: '5', text: '5'},
                    {key: '6', text: '6'},
                  ],
                  selectedKey : '3'
                }),
                PropertyPaneDropdown('graphCalParams.rangeEnd', {
                  label : 'Number of months after today',
                  options: [
                    {key: '1', text: '1'},
                    {key: '2', text: '2'},
                    {key: '3', text: '3'},
                    {key: '4', text: '4'},
                    {key: '5', text: '5'},
                    {key: '6', text: '6'},
                    {key: '7', text: '7'},
                    {key: '8', text: '8'},
                    {key: '9', text: '9'},
                    {key: '10', text: '10'},
                    {key: '11', text: '11'},
                    {key: '12', text: '12'},
                  ],
                  selectedKey : '4'
                }),
                PropertyPaneDropdown('graphCalParams.pageSize', {
                  label : 'Number of events',
                  options: [
                    {key: '50', text: '50'},
                    {key: '100', text: '100'},
                    {key: '150', text: '150'},
                    {key: '200', text: '200'},
                    {key: '250', text: '250'},
                    {key: '300', text: '300'},
                    {key: '350', text: '350'},
                  ],
                  selectedKey : '100'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
