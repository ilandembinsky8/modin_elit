import { SharePointTreeService } from '../services/SharePointTreeService';
import { DocumentCreationService } from '../services/DocumentCreationService';
import { UserProfileService } from '../services/UserProfileService';

export interface IAddFileProps {
  treeService: SharePointTreeService;
  documentService: DocumentCreationService;
  userProfileService: UserProfileService;
  rootSiteUrl: string;
}
