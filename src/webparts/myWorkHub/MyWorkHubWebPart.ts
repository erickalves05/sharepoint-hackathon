import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'MyWorkHubWebPartStrings';
import MyWorkHub from './components/MyWorkHub';
import { IMyWorkHubProps } from './components/IMyWorkHubProps';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import { createGraphBetaCallFromClient } from './services/GraphBetaService';

export interface IMyWorkHubWebPartProps {
  description: string;
}

import type { ServiceScope } from '@microsoft/sp-core-library';

/** Context with serviceScope (SPFx runtime provides this). */
interface IContextWithServiceScope {
  serviceScope?: ServiceScope;
}

export default class MyWorkHubWebPart extends BaseClientSideWebPart<IMyWorkHubWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _msGraphClient: MSGraphClientV3 | undefined;
  private _callGraphBeta: IMyWorkHubProps['callGraphBeta'];
  private _errorMessage: string | null = null;

  public render(): void {
    const element: React.ReactElement<IMyWorkHubProps> = React.createElement(
      MyWorkHub,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks?.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        msGraphClient: this._msGraphClient,
        callGraphBeta: this._callGraphBeta,
        errorMessage: this._errorMessage ?? undefined,
        onError: (msg: string) => {
          this._errorMessage = msg;
          this.render();
        },
        dismissError: () => {
          this._errorMessage = null;
          this.render();
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    const ctx = this.context as unknown as IContextWithServiceScope;
    const serviceScope = ctx.serviceScope;
    if (!serviceScope) {
      return this._getEnvironmentMessage().then(message => {
        this._environmentMessage = message;
      });
    }

    const graphFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);

    return Promise.all([
      graphFactory.getClient('3'),
      this._getEnvironmentMessage()
    ]).then(([client, message]) => {
      this._msGraphClient = client;
      this._callGraphBeta = createGraphBetaCallFromClient(client);
      this._environmentMessage = message;
      this.render();
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
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
