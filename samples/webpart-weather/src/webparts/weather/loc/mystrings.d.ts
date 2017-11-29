declare interface IWeatherWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  LocationFieldLabel: string;
}

declare module 'WeatherWebPartStrings' {
  const strings: IWeatherWebPartStrings;
  export = strings;
}
