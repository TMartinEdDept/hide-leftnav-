import { override } from '@microsoft/decorators';  
import { Log } from '@microsoft/sp-core-library';  
import {  
  BaseApplicationCustomizer  
} from '@microsoft/sp-application-base';  
  
import * as strings from 'HideLeftNavApplicationCustomizerStrings';  
  
const LOG_SOURCE: string = 'HideLeftNavApplicationCustomizer';  
  
export interface IHideLeftNavApplicationCustomizerProperties {  
  testMessage: string;  
}  
  
export default class HideLeftNavApplicationCustomizer  
  extends BaseApplicationCustomizer<IHideLeftNavApplicationCustomizerProperties> {  
  
  @override  
  public onInit(): Promise<void> {  
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);  
  
    const css: string = `  
      /* ---------- Left navigation ---------- */  
      #spLeftNav,  
      [data-automationid="SiteNav"],  
      [data-automation-id="SiteNav"],  
      div[data-automationid='leftNav'],  
      div[data-automationid='SiteHeader'] ~ div [role='navigation'] {  
        display: none !important;  
      }  
  
      /* ---------- SharePoint App Bar (icon rail) ---------- */  
      div[class*="sharePointAppBar"],  
      div[data-id="sharepoint-app-bar"],  
      #sp-appBar,  
      #spo-app-bar,  
      #spo-app-bar-pusher,  
      div[data-sp-feature-tag="App Bar"] {  
        display: none !important;  
      }  
  
      /* ---------- Reclaim the empty space ---------- */  
      [data-automationid="contentScrollRegion"],  
      [data-automation-id="contentScrollRegion"] {  
        margin-left: 0 !important;  
        padding-left: 0 !important;  
      }  
  
      #spPageCanvasContent,  
      .CanvasComponent {  
        margin-left: 0 !important;  
        max-width: 100% !important;  
      }  
    `;  
  
    const head: HTMLHeadElement | null =  
      document.head || document.getElementsByTagName('head')[0];  
  
    const style: HTMLStyleElement = document.createElement('style');  
    style.setAttribute('type', 'text/css');  
    style.appendChild(document.createTextNode(css));  
    head.appendChild(style);  
  
    return Promise.resolve();  
  }  
}  