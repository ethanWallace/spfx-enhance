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

export default class ScriptEnhancementApplicationCustomizer
  extends BaseApplicationCustomizer<IScriptEnhancementApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {

    this._addEnhancements();

    return Promise.resolve();
  }

  private async _getScripts(): Promise<ScriptList> {

    const res = await this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/GetByTitle('ScriptEnhancement')/Items",
      SPHttpClient.configurations.v1);

      return await res.json();
  }

  private _addEnhancements(): void {
    this._getScripts()
      .then((response) => {
        if(response != null){
          response.value.forEach((item: ScriptListItem) => {
            console.log(item);
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
}
