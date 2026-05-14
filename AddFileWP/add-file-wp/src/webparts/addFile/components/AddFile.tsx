import * as React from 'react';
import type { IAddFileProps } from './IAddFileProps';
import { DocumentCreator } from './DocumentCreator';

const AddFile: React.FC<IAddFileProps> = (props) => {
  const { treeService, documentService, userProfileService, rootSiteUrl } = props;

  // Extract site title from URL (last segment)
  const urlParts = rootSiteUrl.replace(/\/+$/, '').split('/');
  const rootSiteTitle = urlParts[urlParts.length - 1] || 'Root Site';

  return (
    <DocumentCreator
      treeService={treeService}
      documentService={documentService}
      userProfileService={userProfileService}
      rootSiteUrl={rootSiteUrl}
      rootSiteTitle={rootSiteTitle}
    />
  );
};

export default AddFile;
