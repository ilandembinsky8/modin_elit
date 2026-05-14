import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IUserProfile } from './UserProfileService';

export class DocumentCreationService {
  private _spHttpClient: SPHttpClient;

  constructor(spHttpClient: SPHttpClient) {
    this._spHttpClient = spHttpClient;
  }

  /**
   * Create a text document in the specified folder.
   * Returns the server-relative URL of the created file.
   */
  public async createDocument(
    siteUrl: string,
    folderServerRelativeUrl: string,
    fileName: string,
    userProfile: IUserProfile
  ): Promise<string> {
    // Sanitize file name
    const sanitizedName = this._sanitizeFileName(fileName);

    // Build file content
    const now = new Date();
    const content = [
      `Created by: ${userProfile.displayName}`,
      `Email: ${userProfile.mail || userProfile.userPrincipalName}`,
      `Department: ${userProfile.department}`,
      `Created at: ${now.toLocaleDateString()} ${now.toLocaleTimeString()}`,
      '',
      'This is a generated document.'
    ].join('\r\n');

    // Check if file already exists
    const fileRelativeUrl = `${folderServerRelativeUrl}/${sanitizedName}`;
    const exists = await this._fileExists(siteUrl, fileRelativeUrl);
    if (exists) {
      throw new Error(`A file named "${sanitizedName}" already exists in this folder.`);
    }

    // Create the file using AddUsingPath (handles special characters including Hebrew)
    const encodedFolder = encodeURIComponent(folderServerRelativeUrl);
    const encodedFileName = encodeURIComponent(sanitizedName);
    const apiUrl = `${siteUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${encodedFolder}')/Files/AddUsingPath(decodedurl='${encodedFileName}',overwrite=false)`;

    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });

    const response: SPHttpClientResponse = await this._spHttpClient.post(
      apiUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata'
        },
        body: blob
      }
    );

    if (!response.ok) {
      const errorData = await response.json().catch(() => null);
      const errorMessage = errorData?.['odata.error']?.message?.value || response.statusText;

      if (response.status === 403) {
        throw new Error('You do not have permission to create files in this folder.');
      }

      throw new Error(`Failed to create document: ${errorMessage}`);
    }

    const result = await response.json();
    return result.ServerRelativeUrl || fileRelativeUrl;
  }

  private async _fileExists(siteUrl: string, fileServerRelativeUrl: string): Promise<boolean> {
    const encodedPath = encodeURIComponent(fileServerRelativeUrl);
    const apiUrl = `${siteUrl}/_api/web/GetFileByServerRelativePath(decodedurl='${encodedPath}')/Exists`;

    try {
      const response: SPHttpClientResponse = await this._spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1,
        { headers: { 'Accept': 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        return false;
      }

      const data = await response.json();
      return data.value === true;
    } catch {
      return false;
    }
  }

  /**
   * Sanitize file name: remove invalid characters, ensure .txt extension
   */
  private _sanitizeFileName(name: string): string {
    // Remove characters invalid in SharePoint file names
    let sanitized = name.replace(/[~"#%&*:<>?/\\{|}]/g, '');
    // Remove leading/trailing dots and spaces
    sanitized = sanitized.replace(/^[\s.]+|[\s.]+$/g, '').trim();

    if (!sanitized) {
      sanitized = 'Document';
    }

    // If no extension provided, append .txt
    if (sanitized.indexOf('.') === -1) {
      sanitized += '.txt';
    }

    return sanitized;
  }
}
