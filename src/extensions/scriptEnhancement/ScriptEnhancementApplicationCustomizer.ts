import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { SPHttpClientResponse, SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';

const LOG_SOURCE: string = 'ScriptEnhancementApplicationCustomizer';

export interface IScriptEnhancementApplicationCustomizerProperties {
}

export interface ScriptList {
  value: ScriptListItem[];
}

export interface ScriptListItem {
  Title: string;
  Content: string;
  ContentType0: string;
  isActive: Boolean;
}

export interface IHubSiteData {
  logoUrl: string;
  name: string;
  navigation: any[];
  themeKey: string;
  url: string;
  usesMetadataNavigation: boolean;
}

export interface IHubSiteDataResponse {
  '@odata.null'?: boolean;
  value?: string;
}

export default class ScriptEnhancementApplicationCustomizer
  extends BaseApplicationCustomizer<IScriptEnhancementApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {

    this.context.application.navigatedEvent.add(this, this._render);

    return Promise.resolve();
  }

  private _getConnectedHubSiteData(): Promise<IHubSiteData> {

    return new Promise<IHubSiteData>((resolve: (connectedHubSiteData: IHubSiteData) => void, reject: (error: any) => void): void => {
      const headers: Headers = new Headers();
      headers.append('accept', 'application/json;odata.metadata=none');

      this.context.spHttpClient
        .get(`${this.context.pageContext.web.absoluteUrl}/_api/web/hubsitedata`, SPHttpClient.configurations.v1, {
          headers: headers
        })
        .then((res: SPHttpClientResponse): Promise<IHubSiteDataResponse> => {
          return res.json();
        })
        .then((res: IHubSiteDataResponse): void => {
          // the site is not connected to a hub site and is not a hub site itself
          if (res['@odata.null'] === true) {
            resolve(undefined);
            return;
          }

          try {
            // parse the hub site data from the value property to JSON
            const hubSiteData: IHubSiteData = JSON.parse(res.value);
            resolve(hubSiteData);
          } catch (e) {
            reject(e);
          }
        })
        .catch((error): void => {
          reject(error);
        });
    });
  }

  private async _getScripts(hubSiteUrl: string): Promise<ScriptList> {

    if(hubSiteUrl !== "") {
      const res = await this.context.spHttpClient.get(
        hubSiteUrl +
        "/_api/web/lists/GetByTitle('ScriptEnhancement')/Items",
        SPHttpClient.configurations.v1);

        return await res.json();
    } else {
      const res = await this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl +
        "/_api/web/lists/GetByTitle('ScriptEnhancement')/Items",
        SPHttpClient.configurations.v1);

        return await res.json();
    }
  }

  private _addEnhancements(hubSiteUrl: string): void {
    this._getScripts(hubSiteUrl)
      .then((response) => {
        if(response != null){
          response.value.forEach((item: ScriptListItem) => {

            if(item.ContentType0 == 'Script' && item.isActive){

              let script = document.createElement('script');
              script.innerText = item.Content;
              document.getElementsByTagName('head')[0].appendChild(script);

            } else if(item.ContentType0 == "Style" && item.isActive) {

              let style = document.createElement('style');
              style.innerText = item.Content;
              document.getElementsByTagName('head')[0].appendChild(style);

            }

          });
        }
      });
  }

  private _render(): void {
    this._getConnectedHubSiteData()
      .then((connectedHubSiteData: IHubSiteData) => {
        if(connectedHubSiteData != undefined) {
          return this._addEnhancements(connectedHubSiteData.url);
        } else {
          const siteUrl = "";
          return this._addEnhancements(siteUrl);
        }
      })
  }
}
