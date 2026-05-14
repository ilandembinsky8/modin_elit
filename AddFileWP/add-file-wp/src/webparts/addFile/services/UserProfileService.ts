import { MSGraphClientV3 } from '@microsoft/sp-http';

export interface IUserProfile {
  displayName: string;
  mail: string;
  userPrincipalName: string;
  department: string;
}

export class UserProfileService {
  private _graphClient: MSGraphClientV3;

  constructor(graphClient: MSGraphClientV3) {
    this._graphClient = graphClient;
  }

  public async getCurrentUser(): Promise<IUserProfile> {
    const user = await this._graphClient
      .api('/me')
      .select('displayName,mail,userPrincipalName,department')
      .get();

    return {
      displayName: user.displayName || '',
      mail: user.mail || user.userPrincipalName || '',
      userPrincipalName: user.userPrincipalName || '',
      department: user.department || ''
    };
  }
}
