/*
?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"b3df0dbe-b759-431a-a81a-ebea6392ba9e":{"location":"ClientSideExtension.ApplicationCustomizer","properties":{"testMessage":"Hello as property!"}}}
*/

import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IODataWeb } from '@microsoft/sp-odata-types';
import * as jQuery from 'jquery';
import * as strings from 'HeaderMenuApplicationCustomizerStrings';
import styles from './HeaderMenu.module.scss';
import { escape } from '@microsoft/sp-lodash-subset'; 

const LOG_SOURCE: string = 'HeaderMenuApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IHeaderMenuApplicationCustomizerProperties {
  // This is an example; replace with your own property
  Top: string
}

export interface TopLinkItem {
  "LinkURL": {
    "Description": "",
    "Url": ""
  }
}

export interface TopLinksList {
  value: TopLinkItem[];
}



/** A Custom Action which can be run during execution of a Client Side Application */
export default class HeaderMenuApplicationCustomizer
  extends BaseApplicationCustomizer<IHeaderMenuApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {

    this._renderHeader();

    return Promise.resolve();
  }

  private _renderHeader(): void {
    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });
  
      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }
  
      if (this.properties) {
  
        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
            <div class="${styles.top}">
              <button id="topMenuToggle"><i class="ms-Icon ms-Icon--CollapseMenu" aria-hidden="true"></i> Toggle</button>
              <div id="topMenu" class="${styles.topMenu}"></div>
            </div>
          `;
          this._renderListAsync();
        }
      }
    }
  }

  private _renderListAsync(): void {
    this._getListData()
      .then((response) => {
        console.log(response.value);
        this._renderList(response.value);
      });
    
  }

  private _renderList(items: TopLinkItem[]): void {
    let html: string = '<ul>';
    items.forEach((item: TopLinkItem) => {
      console.log(item.LinkURL.Description);
      html += `
      <li><a href="${item.LinkURL.Url}">${item.LinkURL.Description}</a></li>
      `;
    });
    html += '</ul>';

    const listContainer: Element = this._topPlaceholder.domElement.querySelector('#topMenu');
    listContainer.innerHTML = html;

    this._setButtonAction();
  }

  private _setButtonAction(): void {
    // is there a benefit to doing this vs. just using jQuery selector?
    const topMenuButton: Element = this._topPlaceholder.domElement.querySelector('#topMenuToggle');
    const topMenuDiv: Element = this._topPlaceholder.domElement.querySelector("#topMenu");
    if (topMenuButton && topMenuDiv){
      jQuery(topMenuButton).click(function(e){
        e.preventDefault();
        e.stopPropagation();
        if (jQuery(topMenuDiv).css("display") == "none"){
          jQuery(topMenuDiv).css("display","block");
        } else {
          jQuery(topMenuDiv).css("display","none");
        }
      });
    }
    
  }

  private _getListData(): Promise<TopLinksList> {
    
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getByTitle('TopLinks')/items?$select=LinkURL&$orderBy=LinkOrder%20asc`, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {
      return response.json();
    });
  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top placeholder.');
  }
}
