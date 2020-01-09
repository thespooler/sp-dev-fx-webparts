import * as React from 'react';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { Link } from 'office-ui-fabric-react/lib/Link';

import { ICalendarEventsProps } from './ICalendarEventProps';

import * as strings from 'ControlStrings';

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { useState, useEffect } from 'react';
import { CalendarEventCard, IUser } from '../../..';

import styles from './../../../controls/calendareventcard/CalendarEventCard.module.scss';

interface ICalendarEvent {
  id: string;
  start: { 
    dateTime: string; 
    timeZone: string; 
  };
  subject:string;
  attendees: {
    type: string;
    status: {
      response: string;
      time: string;
    };
    emailAddress: {
      name: string;
      address: string;
    };
  }[];
}

const getCalendarEvents: (context: WebPartContext, calendarGroupId: string, numberUpcomingDays: number, eventCategory: string) => Promise<IUser[]> = (context: WebPartContext, calendarGroupId: string, numberUpcomingDays: number, eventCategory: string) => {
  const today = new Date();
  const nextYear = new Date();
  nextYear.setFullYear(today.getFullYear() + 1);
  let filter = !!eventCategory ? `&$filter=categories/any(c:c+eq+'${eventCategory}')` : "";
  const url = `https://graph.microsoft.com/v1.0/groups/${calendarGroupId}/calendarView?startdatetime=${today.toLocaleDateString("en-CA")}T00:00:00.0000Z&enddatetime=${nextYear.toLocaleDateString("en-CA")}T00:00:00.000Z&$top=${numberUpcomingDays}&$orderby=start/datetime&$select=attendees,start,subject${filter}`;

  return context.msGraphClientFactory.getClient().then(graph => graph.api(url).get().then((events: { value: ICalendarEvent[] }) => events.value.map(e => ({
      eventDate: e.start.dateTime,
      eventTitle: e.subject,
      userEmail: e.attendees[0].emailAddress.address,
      userName: e.attendees[0].emailAddress.name,
      key: e.attendees[0].emailAddress.address
    } as IUser))));
};

export const CalendarEvents: React.FunctionComponent<ICalendarEventsProps> = (props: ICalendarEventsProps) => {
  const [users, setUsers] = useState([]);
  useEffect(() => {
    getCalendarEvents(props.context, props.calendarGroupId, props.numberUpcomingDays, props.calendarEventCategory).then(u => setUsers(u));
  }, [props.calendarGroupId, props.numberUpcomingDays, props.calendarEventCategory]);
  return <>
        <WebPartTitle displayMode={props.displayMode}
          title={props.title}
          updateProperty={props.updateProperty}
          moreLink={ <Link href={ `https://outlook.office.com/calendar/group/${new URL(props.context.pageContext.site.absoluteUrl).host.replace('.sharepoint.', '.onmicrosoft.')}/${props.calendarGroupMailNickname}/view/month` }>{strings.ShowCalendar}</Link> }
        />
        {
          users.length === 0 ?
            <Placeholder iconName="Calendar"
              iconText={strings.MessageNoEvent}
              description={strings.MessageNoEvent}
              hideButton={true} />
          :
            <div className={styles.calendarEvent}>
            {
              users.map((user: IUser) => <CalendarEventCard {...user} imageUrl={props.imageUrl} />)
            }
            </div>
        }
      </>;
};

export default CalendarEvents;
