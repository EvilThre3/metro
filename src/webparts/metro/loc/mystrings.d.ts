declare interface IMetroWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'MetroWebPartStrings' {
  const strings: IMetroWebPartStrings;
  export = strings;
}
