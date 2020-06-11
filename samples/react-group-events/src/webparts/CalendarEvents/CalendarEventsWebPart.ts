import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import {
  IPropertyPaneChoiceGroupOption,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown
} from "@microsoft/sp-property-pane";
import { PropertyFieldSitePicker, PropertyFieldListPicker, IPropertyFieldSite } from '@pnp/spfx-property-controls';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from 'CalendarEventsWebPartStrings';
import CalendarEvents from './components/CalendarEvents';
import { ICalendarEventsProps, EventSourceType } from './components/ICalendarEventProps';

import { graph } from "@pnp/graph/presets/all";
import { sp } from '@pnp/pnpjs';
import { personaPresenceSize } from 'office-ui-fabric-react/lib/Persona';

export interface ICalendarEventWebPartProps {
  title: string;
  numberUpcomingEvents: number;
  imageIndex: any;
  eventSourceType: EventSourceType;
  calendarGroup: string;
  calendarEventCategory: string;
  siteEventSource: IPropertyFieldSite[];
  listEventSource: string;
}

interface IMSGraphGroup {
  id: string;
  displayName: string;
  mailNickname: string;
  visibility: "Public" | "Private" | "Hiddenmembership" | null | undefined;
}

export const ProvidedImages: string[] = [ require('../../../assets/award.svg'), require('../../../assets/cake.svg') ];

function extractTitleFromPath(path: string): string {
  const fileonly = path.substring(path.lastIndexOf('/') + 1);
  // Require + set-webpack-public-path-plugin + ClienSiteAssets = file_uniqueid.ext
  return fileonly.substring(0,1).toUpperCase() + fileonly.substring(1, fileonly.indexOf('_'));
}

export default class CalendarEventsWebPart extends BaseClientSideWebPart<ICalendarEventWebPartProps> {

  private availableCalendarGroups: IPropertyPaneDropdownOption[] = undefined;

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup(this.context);
      graph.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const [calendarGroupId, calendarGroupMailNickname] = (this.properties.calendarGroup || '/').split('/');
    const element: React.ReactElement<ICalendarEventsProps> = React.createElement(
      CalendarEvents,
      {
        title: this.properties.title,
        numberUpcomingDays: this.properties.numberUpcomingEvents,
        context: this.context,
        displayMode: this.displayMode,
        imageUrl: ProvidedImages[this.properties.imageIndex],
        calendarGroupId,
        calendarGroupMailNickname,
        calendarEventCategory: this.properties.calendarEventCategory,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        eventSourceType: this.properties.eventSourceType,
        siteEventSource: this.properties.siteEventSource,
        listEventSource: this.properties.listEventSource,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.1');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    let dataGroup = [];
    
    if (this.properties.eventSourceType == "GroupCalendar")
    {
      if (this.availableCalendarGroups === undefined) {
        this.context.msGraphClientFactory.getClient().then(client => client
          .api("https://graph.microsoft.com/v1.0/groups?$select=id,displayName,mailNickname,visibility")
          .get().then((groups: { value: IMSGraphGroup[] }) => {
            this.availableCalendarGroups = groups.value
              .filter(g => g.visibility === "Public")
              .map(g => ({ key: `${g.id}/${g.mailNickname}`, text: g.displayName } as IPropertyPaneDropdownOption));
            this.context.propertyPane.refresh();
          }));
      }

      dataGroup = [
        PropertyPaneDropdown("calendarGroup", {
          label: strings.CalendarGroupId,
          options: this.availableCalendarGroups,
          selectedKey: this.properties.calendarGroup
        }),
        PropertyPaneDropdown("calendarEventCategory", {
          label: strings.CalendarEventCategory,
          options: ["", "Purple category", "Blue category", "Green category", "Yellow category", "Orange category", "Red category"]
            .map(c => ({key: c, text: c} as IPropertyPaneDropdownOption)),
          selectedKey: this.properties.calendarEventCategory,
        })
      ];
    }
    else if (this.properties.eventSourceType == "SPList")
    {
      dataGroup = [
        PropertyFieldSitePicker("siteEventSource", {
          label: strings.DataSourceTypeSelectSite,
          context: this.context,
          multiSelect: false,
          initialSites: this.properties.siteEventSource,
          onPropertyChange: this.onPropertyPaneFieldChanged,
          properties: this.properties,
          key: "siteFieldId"
        })
      ];

      if (this.properties.siteEventSource) {
        dataGroup.push(
          PropertyFieldListPicker("listEventSource", {
            label: strings.DataSourceTypeSelectList,
            context: this.context,
            onPropertyChange: this.onPropertyPaneFieldChanged,
            multiSelect: false,
            properties: this.properties,
            key: "listFieldId",
            includeHidden: false,
            selectedList: this.properties.listEventSource,
            webAbsoluteUrl: this.properties.siteEventSource[0].url
        }));
      }
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneChoiceGroup("eventSourceType", {
                  label: strings.DataSourceType,
                  options: [
                    {
                      key: "SPList",
                      text: strings.DataSourceTypeList
                    },
                    {
                      key: "GroupCalendar",
                      text: strings.DataSourceTypeGroup
                    }
                  ]
                }),
                ...dataGroup,
                PropertyFieldNumber("numberUpcomingEvents", {
                  key: "numberUpcomingEvents",
                  label: strings.NumberUpComingEventsLabel,
                  description: strings.NumberUpComingEventsLabel,
                  value: this.properties.numberUpcomingEvents,
                  maxValue: 364,
                  minValue: 1,
                  disabled: false
                }),
                PropertyPaneChoiceGroup('imageIndex', {
                  label: strings.BackgroundImage,
                  options: ProvidedImages.map((imageUrl, i) => ({
                      text: extractTitleFromPath(imageUrl), 
                      key: i,
                      checked: i == this.properties.imageIndex,
                      imageSize: { width: 80, height: 80 },
                      imageSrc: imageUrl,
                      selectedImageSrc: imageUrl
                    } as IPropertyPaneChoiceGroupOption)
                  )
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
