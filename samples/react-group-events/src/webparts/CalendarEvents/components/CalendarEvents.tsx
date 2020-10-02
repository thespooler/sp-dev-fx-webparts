import * as React from 'react';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { Link } from 'office-ui-fabric-react/lib/Link';

import { ICalendarEventsProps } from './ICalendarEventProps';

import * as strings from 'ControlStrings';

import { useState, useEffect } from 'react';
import { CalendarEventCard, IUserEvent } from '../../..';

import { Web } from '@pnp/sp';

import styles from './../../../controls/calendareventcard/CalendarEventCard.module.scss';
import { sp } from '@pnp/pnpjs';
import * as moment from 'moment';

import { groupBy, sortBy } from '@microsoft/sp-lodash-subset';
import { Promise } from 'es6-promise';

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

const getCalendarEvents: (props: ICalendarEventsProps) => Promise<IUserEvent[]> = (props) => {
  let { context, eventSourceType, calendarGroupId, numberUpcomingDays, calendarEventCategory } = props;
  const today = new Date();
  const nextYear = new Date();
  nextYear.setFullYear(today.getFullYear() + 1);

  if (eventSourceType == "SPList" && props.siteEventSource) {
    const untilDay = new Date();
    untilDay.setDate(untilDay.getDate() + props.numberUpcomingDays);
    untilDay.setFullYear(today.getFullYear());
    const wrapsAround = untilDay < today;

    // (month > today.month || (month == today.month && day >= today.day)) && (month < until.month || (month == until.month && day <= until.day))

    return new Web(props.siteEventSource[0].url)
      .lists.getById(props.listEventSource)
      .items
      .filter(`(month gt ${today.getMonth() + 1} or (month eq ${today.getMonth() + 1} and day ge ${today.getDate() })) ${ wrapsAround ? "or": "and"} (month lt ${untilDay.getMonth() + 1} or (month eq ${untilDay.getMonth() + 1} and day le ${untilDay.getDate() }))`)
      .expand("person")
      .select("Title,id,email,month,day,userAADGUID,person/Title,person/ID,person/Name,person/FirstName,person/LastName,person/UserName,person/EMail,person/Department,person/JobTitle,person/Office")
      .orderBy("month,day")
      .get().then(values => {

        if (wrapsAround)
        {
          let pivotWrap = groupBy(values, p => p.month < today.getMonth() + 1 || (p.month == today.getMonth() + 1 && p.day < today.getDate()));
          pivotWrap["true"] = sortBy(pivotWrap["true"], ['month', 'day']);
          pivotWrap["false"] = sortBy(pivotWrap["false"], ['month', 'day']);
          values = [...pivotWrap["false"], ...pivotWrap["true"]];
        }
        
        return values.map(value => { 
          const date = new Date();
          date.setMonth(value.month - 1);
          date.setDate(value.day);
          return ({
            eventDate: moment(date).format("YYYY-MM-DD"),
            eventTitle: value.Title || (value.person ? value.person.Title : ""),
            key: value.id,
            userEmail: value.email || (value.person ? value.person.EMail : ""),
            userName: (value.person ? value.person.UserName : ""),
            jobDescription: (value.person ? value.person.JobTitle : ""),
          } as IUserEvent);
        });
      }) as Promise<IUserEvent[]>;
  } else if (eventSourceType == "GroupCalendar" && calendarGroupId) {
    let filter = !!calendarEventCategory ? `&$filter=categories/any(c:c+eq+'${calendarEventCategory}')` : "";
    const url = `https://graph.microsoft.com/v1.0/groups/${calendarGroupId}/calendarView?startdatetime=${today.toLocaleDateString("en-CA")}T00:00:00.0000Z&enddatetime=${nextYear.toLocaleDateString("en-CA")}T00:00:00.000Z&$top=${numberUpcomingDays}&$orderby=start/datetime&$select=attendees,start,subject${filter}`;

    return context.msGraphClientFactory.getClient().then(graph => graph.api(url).get().then((events: { value: ICalendarEvent[] }) => events.value.map(e => ({
        eventDate: e.start.dateTime,
        eventTitle: e.subject,
        userEmail: e.attendees[0].emailAddress.address,
        userName: e.attendees[0].emailAddress.name,
        key: e.attendees[0].emailAddress.address
      } as IUserEvent)))) as Promise<IUserEvent[]>;
  } else {
    return Promise.resolve([] as IUserEvent[]);
  }
};

export const CalendarEvents: React.FunctionComponent<ICalendarEventsProps> = (props: ICalendarEventsProps) => {
  const [userEvents, setUserEvents] = useState<IUserEvent[]>([]);
  const [eventsUrl, setEventsUrl] = useState('');
  useEffect(() => {
    getCalendarEvents(props).then(u => setUserEvents(u));
  }, [props.calendarGroupId, props.numberUpcomingDays, props.calendarEventCategory]);

  useEffect(() => {
    if (props.eventSourceType == "GroupCalendar")
    {
      setEventsUrl(`https://outlook.office.com/calendar/group/${new URL(props.context.pageContext.site.absoluteUrl).host.replace('.sharepoint.', '.onmicrosoft.')}/${props.calendarGroupMailNickname}/view/month`);
    }
    else if (props.eventSourceType == "SPList" && props.siteEventSource) {
      new Web(props.siteEventSource[0].url)
        .lists.getById(props.listEventSource)
        .defaultView.get()
        .then(value => setEventsUrl(value.ServerRelativeUrl));
    }
  }, [props.eventSourceType, props.context, props.siteEventSource, props.listEventSource]);

  return <>
        <WebPartTitle displayMode={props.displayMode}
          title={props.title}
          updateProperty={props.updateProperty}
          moreLink={ <Link target="_blank" href={ eventsUrl }>{strings.ShowEvents}</Link> }
        />
        {
          userEvents.length === 0 ?
            <Placeholder iconName="Calendar"
              iconText={strings.MessageNoEvent}
              description={strings.MessageNoEvent}
              hideButton={true} />
          :
            <div className={styles.calendarEvent}>
            {
              userEvents.map((userEvent: IUserEvent) => <CalendarEventCard {...userEvent} imageUrl={props.imageUrl} />)
            }
            </div>
        }
      </>;
};

export default CalendarEvents;
