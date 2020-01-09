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
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from 'CalendarEventsWebPartStrings';
import CalendarEvents from './components/CalendarEvents';
import { ICalendarEventsProps } from './components/ICalendarEventProps';

export interface ICalendarEventWebPartProps {
  title: string;
  numberUpcomingDays: number;
  imageIndex: any;
  calendarGroup: string;
  calendarEventCategory: string;
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
      // other init code may be present
    });
  }

  public render(): void {
    const [calendarGroupId, calendarGroupMailNickname] = (this.properties.calendarGroup || '/').split('/');
    const element: React.ReactElement<ICalendarEventsProps> = React.createElement(
      CalendarEvents,
      {
        title: this.properties.title,
        numberUpcomingDays: this.properties.numberUpcomingDays,
        context: this.context,
        displayMode: this.displayMode,
        imageUrl: ProvidedImages[this.properties.imageIndex],
        calendarGroupId,
        calendarGroupMailNickname,
        calendarEventCategory: this.properties.calendarEventCategory,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
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
                }),
                PropertyFieldNumber("numberUpcomingDays", {
                  key: "numberUpcomingDays",
                  label: strings.NumberUpComingDaysLabel,
                  description: strings.NumberUpComingDaysLabel,
                  value: this.properties.numberUpcomingDays,
                  maxValue: 10,
                  minValue: 5,
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
