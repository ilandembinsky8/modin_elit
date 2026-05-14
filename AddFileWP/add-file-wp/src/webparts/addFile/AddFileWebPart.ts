import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'AddFileWebPartStrings';
import AddFile from './components/AddFile';
import { IAddFileProps } from './components/IAddFileProps';
import { SharePointTreeService } from './services/SharePointTreeService';
import { DocumentCreationService } from './services/DocumentCreationService';
import { UserProfileService } from './services/UserProfileService';

export interface IAddFileWebPartProps {
  rootSiteUrl: string;
}

export default class AddFileWebPart extends BaseClientSideWebPart<IAddFileWebPartProps> {

  private _treeService!: SharePointTreeService;
  private _documentService!: DocumentCreationService;
  private _userProfileService!: UserProfileService;

  public render(): void {
    const rootSiteUrl = this.properties.rootSiteUrl || 'https://modil.sharepoint.com/sites/MODIL';

    const element: React.ReactElement<IAddFileProps> = React.createElement(
      AddFile,
      {
        treeService: this._treeService,
        documentService: this._documentService,
        userProfileService: this._userProfileService,
        rootSiteUrl: rootSiteUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._treeService = new SharePointTreeService(this.context.spHttpClient);
    this._documentService = new DocumentCreationService(this.context.spHttpClient);

    const graphClient: MSGraphClientV3 = await this.context.msGraphClientFactory.getClient('3');
    this._userProfileService = new UserProfileService(graphClient);
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
                PropertyPaneTextField('rootSiteUrl', {
                  label: 'Root Site URL',
                  description: 'The root site collection URL to browse',
                  value: this.properties.rootSiteUrl || 'https://modil.sharepoint.com/sites/MODIL'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
