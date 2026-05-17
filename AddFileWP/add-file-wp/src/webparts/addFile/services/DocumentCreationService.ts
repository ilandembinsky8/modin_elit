import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IUserProfile } from './UserProfileService';
import { IDocumentType } from '../models/IDocumentType';
import { OfficeFileGenerator } from './OfficeFileGenerator';

export class DocumentCreationService {
  private _spHttpClient: SPHttpClient;
  private _fileGenerator: OfficeFileGenerator;

  constructor(spHttpClient: SPHttpClient) {
    this._spHttpClient = spHttpClient;
    this._fileGenerator = new OfficeFileGenerator();
  }

  /**
   * Create a document in the specified folder.
   * Returns the server-relative URL of the created file.
   */
  public async createDocument(
    siteUrl: string,
    folderServerRelativeUrl: string,
    fileName: string,
    documentType: IDocumentType,
    userProfile: IUserProfile
  ): Promise<string> {
    // Sanitize file name and apply the correct extension
    const sanitizedName = this._sanitizeFileName(fileName, documentType.extension);

    // Check if file already exists
    const fileRelativeUrl = folderServerRelativeUrl + '/' + sanitizedName;
    const exists = await this._fileExists(siteUrl, fileRelativeUrl);
    if (exists) {
      throw new Error('A file named "' + sanitizedName + '" already exists in this folder.');
    }

    // Generate file content based on type
    const now = new Date();
    const dateStr = now.toLocaleDateString() + ' ' + now.toLocaleTimeString();
    let blob: Blob;

    switch (documentType.key) {
      case 'docx':
        blob = await this._fileGenerator.generateDocx(userProfile, dateStr);
        break;
      case 'xlsx':
        blob = await this._fileGenerator.generateXlsx(userProfile, dateStr);
        break;
      case 'pptx':
        blob = await this._fileGenerator.generatePptx(userProfile, dateStr);
        break;
      default:
        blob = this._fileGenerator.generateTxt(userProfile, dateStr);
        break;
    }

    // Upload the file using AddUsingPath (handles special characters including Hebrew)
    const encodedFolder = encodeURIComponent(folderServerRelativeUrl);
    const encodedFileName = encodeURIComponent(sanitizedName);
    const apiUrl = siteUrl + "/_api/web/GetFolderByServerRelativePath(decodedurl='" + encodedFolder + "')/Files/AddUsingPath(decodedurl='" + encodedFileName + "',overwrite=false)";

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
      const errorData = await response.json().catch(function() { return null; });
      let errorMessage: string = response.statusText;
      if (errorData && errorData['odata.error'] && errorData['odata.error'].message) {
        errorMessage = errorData['odata.error'].message.value || errorMessage;
      }

      if (response.status === 403) {
        throw new Error('You do not have permission to create files in this folder.');
      }

      throw new Error('Failed to create document: ' + errorMessage);
    }

    const result = await response.json();
    return result.ServerRelativeUrl || fileRelativeUrl;
  }

  private async _fileExists(siteUrl: string, fileServerRelativeUrl: string): Promise<boolean> {
    const encodedPath = encodeURIComponent(fileServerRelativeUrl);
    const apiUrl = siteUrl + "/_api/web/GetFileByServerRelativePath(decodedurl='" + encodedPath + "')/Exists";

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
   * Sanitize file name: remove invalid characters, apply correct extension
   */
  private _sanitizeFileName(name: string, defaultExtension: string): string {
    // Remove characters invalid in SharePoint file names
    let sanitized = name.replace(/[~"#%&*:<>?/\\{|}]/g, '');
    // Remove leading/trailing dots and spaces
    sanitized = sanitized.replace(/^[\s.]+|[\s.]+$/g, '').trim();

    if (!sanitized) {
      sanitized = 'Document';
    }

    // Remove any existing extension from the name
    const dotIndex = sanitized.lastIndexOf('.');
    if (dotIndex > 0) {
      sanitized = sanitized.substring(0, dotIndex);
    }

    // Append the correct extension for the selected type
    sanitized += defaultExtension;

    return sanitized;
  }
}
