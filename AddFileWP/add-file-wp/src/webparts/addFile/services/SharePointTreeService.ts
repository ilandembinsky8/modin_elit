import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ITreeNode } from '../models/ITreeNode';

export class SharePointTreeService {
  private _spHttpClient: SPHttpClient;

  constructor(spHttpClient: SPHttpClient) {
    this._spHttpClient = spHttpClient;
  }

  /**
   * Get the root site node
   */
  public getRootNode(rootSiteUrl: string, rootSiteTitle: string): ITreeNode {
    return {
      id: rootSiteUrl,
      title: rootSiteTitle,
      type: 'site',
      webUrl: rootSiteUrl,
      children: [],
      isLoaded: false,
      isLoading: false,
      selectable: false
    };
  }

  /**
   * Look up the user's department in the Departments list on the root site.
   * Returns the SitePath value if found, or undefined if not.
   */
  public async getDepartmentSitePath(rootSiteUrl: string, departmentName: string): Promise<string | undefined> {
    if (!departmentName) {
      return undefined;
    }

    const filterValue = departmentName.replace(/'/g, "''");
    const apiUrl = `${rootSiteUrl}/_api/web/lists/getbytitle('Departments')/items?$filter=Title eq '${filterValue}'&$select=Title,SitePath&$top=1`;

    try {
      const response: SPHttpClientResponse = await this._spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        console.warn('Failed to query Departments list:', response.statusText);
        return undefined;
      }

      const data = await response.json();
      if (data.value && data.value.length > 0 && data.value[0].SitePath) {
        return data.value[0].SitePath;
      }

      return undefined;
    } catch (error) {
      console.warn('Error querying Departments list:', error);
      return undefined;
    }
  }

  /**
   * Load children for a site node: subsites + document libraries
   */
  public async loadSiteChildren(siteUrl: string): Promise<ITreeNode[]> {
    const [subsites, libraries] = await Promise.all([
      this._getSubsites(siteUrl),
      this._getDocumentLibraries(siteUrl)
    ]);
    return [...subsites, ...libraries];
  }

  /**
   * Load folders inside a document library root
   */
  public async loadLibraryFolders(siteUrl: string, libraryServerRelativeUrl: string): Promise<ITreeNode[]> {
    return this._getFolders(siteUrl, libraryServerRelativeUrl);
  }

  /**
   * Load subfolders inside a folder
   */
  public async loadFolderChildren(siteUrl: string, folderServerRelativeUrl: string): Promise<ITreeNode[]> {
    return this._getFolders(siteUrl, folderServerRelativeUrl);
  }

  private async _getSubsites(siteUrl: string): Promise<ITreeNode[]> {
    const apiUrl = `${siteUrl}/_api/web/webs?$select=Title,Url,ServerRelativeUrl`;
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      apiUrl,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      console.error(`Failed to load subsites for ${siteUrl}:`, response.statusText);
      return [];
    }

    const data = await response.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (data.value || []).map((web: any) => ({
      id: web.Url,
      title: web.Title,
      type: 'site' as const,
      webUrl: web.Url,
      children: [],
      isLoaded: false,
      isLoading: false,
      selectable: false
    }));
  }

  private async _getDocumentLibraries(siteUrl: string): Promise<ITreeNode[]> {
    const apiUrl = `${siteUrl}/_api/web/lists?$filter=BaseTemplate eq 101 and Hidden eq false&$select=Title,Id,RootFolder/ServerRelativeUrl&$expand=RootFolder`;
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      apiUrl,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      console.error(`Failed to load libraries for ${siteUrl}:`, response.statusText);
      return [];
    }

    const data = await response.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (data.value || []).filter((list: any) => {
      // Filter out common system libraries
      const systemLibraries = ['Style Library', 'Form Templates', 'Site Assets', 'Site Pages', 'appdata', '_catalogs'];
      return !systemLibraries.some(sl => sl.toLowerCase() === (list.Title || '').toLowerCase());
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    }).map((list: any) => ({
      id: `lib_${list.Id}`,
      title: list.Title,
      type: 'library' as const,
      webUrl: siteUrl,
      serverRelativeUrl: list.RootFolder.ServerRelativeUrl,
      children: [],
      isLoaded: false,
      isLoading: false,
      selectable: true
    }));
  }

  private async _getFolders(siteUrl: string, parentFolderServerRelativeUrl: string): Promise<ITreeNode[]> {
    const encodedPath = encodeURIComponent(parentFolderServerRelativeUrl);
    const apiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${encodedPath}')/Folders?$filter=Name ne 'Forms'&$select=Name,ServerRelativeUrl,UniqueId`;
    const response: SPHttpClientResponse = await this._spHttpClient.get(
      apiUrl,
      SPHttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json;odata=nometadata' } }
    );

    if (!response.ok) {
      console.error(`Failed to load folders for ${parentFolderServerRelativeUrl}:`, response.statusText);
      return [];
    }

    const data = await response.json();
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (data.value || []).map((folder: any) => ({
      id: `folder_${folder.UniqueId}`,
      title: folder.Name,
      type: 'folder' as const,
      webUrl: siteUrl,
      serverRelativeUrl: folder.ServerRelativeUrl,
      children: [],
      isLoaded: false,
      isLoading: false,
      selectable: true
    }));
  }
}
