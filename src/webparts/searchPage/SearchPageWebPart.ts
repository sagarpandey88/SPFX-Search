import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SearchPageWebPartStrings';
import SearchPage, { ISearchPageProps } from './components/MainSearchPage/SearchPage';
import { getQueryParam, getHubSiteId } from '../../services/searchUtils';



export interface ISearchPageWebPartProps {
  description: string;
}

export default class SearchPageWebPart extends BaseClientSideWebPart<ISearchPageWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _scopeFilters: string[] = [];
  private _scopeLabel: string | null = null;
  private _initialQuery: string | null = null;

  public render(): void {
    const element: React.ReactElement<ISearchPageProps> = React.createElement(
      SearchPage,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        initialQuery: this._initialQuery,

        //custom properties resolved in onInit and passed down to the SearchPage component, which is responsible for performing searches with the correct scope.
        scopeFilters: this._scopeFilters,
        scopeLabel: this._scopeLabel

      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });

    // Determine search scope once, before the first render. Read querystring
    const sourceSiteId = getQueryParam("sourceSiteId");
    this._initialQuery = getQueryParam("q");

    if (sourceSiteId) {
      // Child site redirect: scope to the specific site collection by GUID.
      this._scopeFilters = [`SiteId:"${sourceSiteId}"`];
      const siteTitle = getQueryParam("sourceSiteTitle") || "a specific site collection";
      this._scopeLabel = `Showing results from ${siteTitle}`;
    } else {
      // Hub site: scope to hub and all associated sites via DepartmentId.
      const hubSiteId = await getHubSiteId(this.context);
      if (hubSiteId) {
        this._scopeFilters = [`DepartmentId:"${hubSiteId}"`];
        this._scopeLabel = "Showing results from all sites";
      }
      // No hub association — scope left empty, tenant-wide results, no label.
    }
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
