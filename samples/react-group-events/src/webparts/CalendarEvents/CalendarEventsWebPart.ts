import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import {
  IPropertyPaneChoiceGroupOption,
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from "@microsoft/sp-property-pane";
import { PropertyFieldSitePicker, PropertyFieldListPicker, IPropertyFieldSite } from '@pnp/spfx-property-controls';
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from 'CalendarEventsWebPartStrings';
import CalendarEvents from './components/CalendarEvents';
import { ICalendarEventsProps, EventSourceType } from './components/ICalendarEventProps';

import { graph } from "@pnp/graph/presets/all";
import { sp } from '@pnp/pnpjs';

import {
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme
} from '@microsoft/sp-component-base';

export interface ICalendarEventWebPartProps {
  title: string;
  numberUpcomingEvents: number;
  imageIndex: any;
  eventSourceType: EventSourceType;
  calendarGroup: string;
  calendarEventCategory: string;
  eventSourceSite: IPropertyFieldSite[];
  eventSourceList: string;
  showEventsTargetUrl: string;
}

interface IMSGraphGroup {
  id: string;
  displayName: string;
  mailNickname: string;
  visibility: "Public" | "Private" | "Hiddenmembership" | null | undefined;
}

export const ProvidedImages: string[] = [ require('../../../assets/award.svg'), require('../../../assets/cake.svg'), require('../../../assets/icn_anniversaire_20200622_v02_mt-02.svg') ];

function extractTitleFromPath(path: string): string {
  const fileonly = path.substring(path.lastIndexOf('/') + 1);
  // Require + set-webpack-public-path-plugin + ClienSiteAssets = file_uniqueid.ext
  return fileonly.substring(0,1).toUpperCase() + fileonly.substring(1, fileonly.indexOf('_'));
}

export default class CalendarEventsWebPart extends BaseClientSideWebPart<ICalendarEventWebPartProps> {

  private availableCalendarGroups: IPropertyPaneDropdownOption[] = undefined;
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  public onInit(): Promise<void> {

    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return super.onInit().then(_ => {
      sp.setup(this.context);
      graph.setup({
        spfxContext: this.context
      });
    });
  }
 
  /**
   * Update the current theme variant reference and re-render.
   *
   * @param args The new theme
   */
  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
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
        siteEventSource: this.properties.eventSourceSite,
        listEventSource: this.properties.eventSourceList,
        showEventsTargetUrl: this.properties.showEventsTargetUrl,
        themeVariant: this._themeVariant
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
        PropertyFieldSitePicker("eventSourceSite", {
          label: strings.DataSourceTypeSelectSite,
          context: this.context,
          multiSelect: false,
          initialSites: this.properties.eventSourceSite,
          onPropertyChange: this.onPropertyPaneFieldChanged,
          properties: this.properties,
          key: "siteFieldId"
        })
      ];

      if (this.properties.eventSourceSite && this.properties.eventSourceSite.length > 0) {
        dataGroup.push(
          PropertyFieldListPicker("eventSourceList", {
            label: strings.DataSourceTypeSelectList,
            context: this.context,
            onPropertyChange: this.onPropertyPaneFieldChanged,
            multiSelect: false,
            properties: this.properties,
            key: "listFieldId",
            includeHidden: false,
            selectedList: this.properties.eventSourceList,
            webAbsoluteUrl: this.properties.eventSourceSite[0].url
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
                }),
                PropertyPaneTextField('showEventsTargetUrl', { 
                  label:strings.ShowEventsTargetUrl, 
                  value: this.properties.showEventsTargetUrl,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
