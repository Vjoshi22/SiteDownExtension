import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as $ from "jquery";
import  SiteDownPage, {IreactSiteProps} from "../components/SiteDownPage";
import * as strings from 'PmoSiteStatusApplicationCustomizerStrings';

var configuration_arr:any = [];
const LOG_SOURCE: string = 'PmoSiteStatusApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPmoSiteStatusApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PmoSiteStatusApplicationCustomizer
  extends BaseApplicationCustomizer<IPmoSiteStatusApplicationCustomizerProperties> {
    private static headerPlaceholder: PlaceholderContent;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    
    this.render();
    return Promise.resolve();
  }
  public onDispose() {
    if (PmoSiteStatusApplicationCustomizer.headerPlaceholder && PmoSiteStatusApplicationCustomizer.headerPlaceholder.domElement) {
      ReactDom.unmountComponentAtNode(PmoSiteStatusApplicationCustomizer.headerPlaceholder.domElement);
    }
  }
  private render() {
    if (this.context.placeholderProvider.placeholderNames.indexOf(PlaceholderName.Top) !== -1) {
      if (!PmoSiteStatusApplicationCustomizer.headerPlaceholder || !PmoSiteStatusApplicationCustomizer.headerPlaceholder.domElement) {
        PmoSiteStatusApplicationCustomizer.headerPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
          onDispose: this.onDispose
        });
      }
      this.loadReactComponent();
    }
    else {
      console.log(`The following placeholder names are available`, this.context.placeholderProvider.placeholderNames);
    }
  }
  /**
   * Start the React rendering of your components
   */
  private loadReactComponent() {
    if (PmoSiteStatusApplicationCustomizer.headerPlaceholder && PmoSiteStatusApplicationCustomizer.headerPlaceholder.domElement) {
      const element: React.ReactElement<IreactSiteProps> = React.createElement(SiteDownPage, {
        context: this.context
      });
      ReactDom.render(element, PmoSiteStatusApplicationCustomizer.headerPlaceholder.domElement);
    }
    else {
      console.log('DOM element of the header is undefined. Start to re-render.');
      this.render();
    }
  }
  
}
