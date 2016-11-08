declare interface INewbidStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'newbidStrings' {
  const strings: INewbidStrings;
  export = strings;
}
