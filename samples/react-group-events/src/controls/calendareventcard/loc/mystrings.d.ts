declare interface IControlStrings {
  Today: string,
  MessageNoEvent: string,
  ShowCalendar: string
}

declare module 'ControlStrings' {
  const strings: IControlStrings;
  export = strings;
}
