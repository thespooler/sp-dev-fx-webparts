declare interface IControlStrings {
  Today: string,
  MessageNoEvent: string,
  ShowEvents: string
}

declare module 'ControlStrings' {
  const strings: IControlStrings;
  export = strings;
}
