import * as React from 'react';
import { ICalendarEventCardProps } from './ICalendarEventCardProps';
import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { Image, ImageFit } from 'office-ui-fabric-react/lib/Image';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Icon } from 'office-ui-fabric-react/lib/Icon';
import * as moment from 'moment';
import {  DocumentCardActions } from 'office-ui-fabric-react/lib/DocumentCard';
import { FunctionComponent } from 'react';

import styles from './CalendarEventCard.module.scss';

const IMG_WIDTH: number = 200;
const IMG_HEIGTH: number = 190;

export const CalendarEventCard: FunctionComponent<ICalendarEventCardProps> = (props: ICalendarEventCardProps) => {
  const isEventToday = moment().isSame(props.eventDate, 'day');
  const photo: string = `/_layouts/15/userphoto.aspx?size=L&username=${props.userEmail}`;
  const formattedDate = moment(props.eventDate).format('Do MMMM');
  const persona = {
    imageUrl: photo ? photo : '',
    imageInitials: props.userName.split(' ').map(a => a.charAt(0).toLocaleUpperCase()).join(''),
    text: props.userName,
    secondaryText: props.jobDescription,
    tertiaryText: formattedDate,
  };

  console.info(props.imageUrl);

  return <div className={[styles.documentCard, isEventToday ? styles.today : ''].join(' ')}>
          <Image
            imageFit={ImageFit.cover}
            src={props.imageUrl}
            width={IMG_WIDTH}
            height={IMG_HEIGTH}
          />
          <Label className={styles.centered}>{props.eventTitle}</Label>
          <Label className={[ styles.centered, styles.eventDate, isEventToday ? styles.eventDateToday : '' ].join(' ')}>{formattedDate}</Label>
          <div className={styles.personaContainer}>
            <Persona
              {...persona}
              size={PersonaSize.regular}
              className={styles.persona}
              onRenderTertiaryText={ personaProps => <div>
                <span className='personaTertiary'>{personaProps.tertiaryText}</span>
              </div> }
            />
          </div>
          <div className={styles.actions}>
            <DocumentCardActions
              actions={[
                {
                  iconProps: { iconName: 'Mail' },
                  onClick: (ev: any) => {
                    ev.preventDefault();
                    ev.stopPropagation();
                  window.location.href = `mailto:${props.userEmail}?subject=${props.eventTitle}!`;
                  },
                  ariaLabel: 'email'
                }
              ]}
            />
            {
            isEventToday && <div className={styles.eventcake}><Icon iconName="GotoToday" /></div>
            }
          </div>
        </div>;
};

export default CalendarEventCard;
