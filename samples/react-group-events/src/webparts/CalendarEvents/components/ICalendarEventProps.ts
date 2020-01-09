import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DisplayMode } from '@microsoft/sp-core-library';

export interface ICalendarEventsProps {
  title: string;
  numberUpcomingDays: number;
  context: WebPartContext;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  calendarGroupId: string;
  calendarGroupMailNickname: string;
  calendarEventCategory: string;
  imageUrl:string;
}
