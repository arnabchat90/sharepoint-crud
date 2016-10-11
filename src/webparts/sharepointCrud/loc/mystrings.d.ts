declare interface ISharepointCrudStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  ListNameFieldLabel: string;
}

declare module 'sharepointCrudStrings' {
  const strings: ISharepointCrudStrings;
  export = strings;
}
