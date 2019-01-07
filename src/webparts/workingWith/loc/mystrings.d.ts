declare interface IWorkingWithWebPartStrings {
  BasicGroupName: string;
  PropertyPaneDescription: string;
  NrOfContactsToShow: string;
  NoContacts: string;
  Loading: string;
  Error: string;
}

declare module 'WorkingWithWebPartStrings' {
  const strings: IWorkingWithWebPartStrings;
  export = strings;
}
