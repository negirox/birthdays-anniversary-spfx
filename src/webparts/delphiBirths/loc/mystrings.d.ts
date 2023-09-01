declare interface IDelphiBirthsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  NumberUpComingDaysLabel: string;
  BirthdayControlDefaultDay: string,
  HappyBirthdayMsg: string,
  NextBirthdayMsg: string,
  HappyAnniversaryMsg: string,
  NextAnniversaryMsg: string,
  MessageNoBirthdays: string
}

declare module 'DelphiBirthsWebPartStrings' {
  const strings: IDelphiBirthsWebPartStrings;
  export = strings;
}
