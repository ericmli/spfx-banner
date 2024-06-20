import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'BannerWebPartStrings';
import Banner from './components/Banner';
import { IBannerProps } from './components/IBannerProps';

export interface IBannerWebPartProps {
  description: string;
  secondsSlick: number;
  checkbox: boolean;
  color: string;
}

export default class BannerWebPart extends BaseClientSideWebPart<IBannerWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IBannerProps> = React.createElement(
      Banner,
      {
        description: this.properties.description,
        secondsSlick: this.properties.secondsSlick,
        checkbox: this.properties.checkbox,
        color: this.properties.color,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        userDisplayEmail: this.context.pageContext.user.email,
        site: this.context.pageContext.site
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {    
      const input = PropertyPaneTextField('description', {
        label: "Name List SharePoint",
      });
      const secondsSlick = PropertyPaneSlider('secondsSlick', {
        label:"Select Seconds for time Carousel",min:2,max:50,value:8,showValue:true,step:1
      });
      const checkbox = PropertyPaneCheckbox('checkbox', {
        text: 'Slick Mode',
        checked: true,
      });
      const inputColor = PropertyPaneTextField('color', {
        label: "Select color in format's CSS for Background",
      });
      return {
        pages: [
          {
            header: {
              description: 'Attention with a list of names in SharePoint, use the correct name or GUID!'
            },
            groups: [
              {
                groupFields: [
                  input,
                  secondsSlick,
                  checkbox,
                  inputColor
                ]
              }
            ]
          }
        ]
      }
  }
}
