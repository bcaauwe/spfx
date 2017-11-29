import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName 
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'AnnouncementsApplicationCustomizerStrings';
import Announcements, { IAnnouncementsProps } from './components/Announcements';

const LOG_SOURCE: string = 'AnnouncementsApplicationCustomizer';

export interface IAnnouncementsApplicationCustomizerProperties {
  siteUrl: string;
  listName: string;
}

export default class AnnouncementsApplicationCustomizer
  extends BaseApplicationCustomizer<IAnnouncementsApplicationCustomizerProperties> {

    private _topPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    if (!this.properties.siteUrl || !this.properties.listName){
      const e: Error = new Error('Missing required configuration parameters');
      console.log(e);
      return Promise.reject(e);
    }

    this._renderPlaceHolders();

    return Promise.resolve<void>();
  }

  private _renderPlaceHolders(): void {
    console.log('Available placeholders: ', this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(', '));

    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });

      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }

      const elem: React.ReactElement<IAnnouncementsProps> = React.createElement(Announcements, {
        siteUrl: this.properties.siteUrl,
        listName: this.properties.listName,
        spHttpClient: this.context.spHttpClient
      });
      ReactDOM.render(elem, this._topPlaceholder.domElement);
    }
  }

  private _onDispose(): void {
    console.log('Disposed custom announcement placeholders');
  }
}
