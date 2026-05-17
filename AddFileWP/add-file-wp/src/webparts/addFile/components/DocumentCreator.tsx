import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Label } from '@fluentui/react/lib/Label';
import { Link } from '@fluentui/react/lib/Link';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { LocationTreePicker } from './LocationTreePicker';
import { ITreeNode } from '../models/ITreeNode';
import { IDocumentType, DocumentTypes } from '../models/IDocumentType';
import { SharePointTreeService } from '../services/SharePointTreeService';
import { DocumentCreationService } from '../services/DocumentCreationService';
import { UserProfileService, IUserProfile } from '../services/UserProfileService';
import styles from './AddFile.module.scss';

export interface IDocumentCreatorProps {
  treeService: SharePointTreeService;
  documentService: DocumentCreationService;
  userProfileService: UserProfileService;
  rootSiteUrl: string;
  rootSiteTitle: string;
}

interface IDocumentCreatorState {
  rootNode: ITreeNode;
  selectedNode: ITreeNode | undefined;
  documentName: string;
  selectedDocType: IDocumentType;
  isCreating: boolean;
  errorMessage: string;
  successMessage: string;
  createdFileUrl: string;
  userProfile: IUserProfile | undefined;
  userProfileError: string;
}

export class DocumentCreator extends React.Component<IDocumentCreatorProps, IDocumentCreatorState> {

  private _docTypeOptions: IDropdownOption[];

  constructor(props: IDocumentCreatorProps) {
    super(props);

    const rootNode = props.treeService.getRootNode(props.rootSiteUrl, props.rootSiteTitle);

    this._docTypeOptions = DocumentTypes.map(function(dt) {
      return {
        key: dt.key,
        text: dt.text,
        data: { icon: dt.icon }
      };
    });

    this.state = {
      rootNode,
      selectedNode: undefined,
      documentName: '',
      selectedDocType: DocumentTypes[0], // Default to Word
      isCreating: false,
      errorMessage: '',
      successMessage: '',
      createdFileUrl: '',
      userProfile: undefined,
      userProfileError: ''
    };
  }

  public async componentDidMount(): Promise<void> {
    try {
      const userProfile = await this.props.userProfileService.getCurrentUser();
      this.setState({ userProfile });

      // Look up department in Departments list to resolve tree root
      if (userProfile.department) {
        const sitePath = await this.props.treeService.getDepartmentSitePath(
          this.props.rootSiteUrl,
          userProfile.department
        );

        if (sitePath) {
          // Build the subsite URL from the SitePath value
          const departmentSiteUrl = this.props.rootSiteUrl.replace(/\/+$/, '') + '/' + sitePath;
          const rootNode = this.props.treeService.getRootNode(departmentSiteUrl, userProfile.department + ' (' + sitePath + ')');
          this.setState({ rootNode });
        }
      }
    } catch (error) {
      this.setState({
        userProfileError: '\u05E9\u05D2\u05D9\u05D0\u05D4 \u05D1\u05D8\u05E2\u05D9\u05E0\u05EA \u05E4\u05E8\u05D5\u05E4\u05D9\u05DC \u05DE\u05E9\u05EA\u05DE\u05E9: ' + (error instanceof Error ? error.message : String(error))
      });
    }
  }

  public render(): React.ReactElement {
    const {
      rootNode,
      selectedNode,
      documentName,
      selectedDocType,
      isCreating,
      errorMessage,
      successMessage,
      createdFileUrl,
      userProfile,
      userProfileError
    } = this.state;

    const canCreate = documentName.trim().length > 0 &&
      selectedNode !== undefined &&
      selectedNode.selectable &&
      !isCreating &&
      userProfile !== undefined;

    // Show file name preview
    let fileNamePreview = '';
    if (documentName.trim()) {
      let cleanName = documentName.trim().replace(/[~"#%&*:<>?/\\{|}]/g, '');
      const dotIdx = cleanName.lastIndexOf('.');
      if (dotIdx > 0) {
        cleanName = cleanName.substring(0, dotIdx);
      }
      if (cleanName) {
        fileNamePreview = cleanName + selectedDocType.extension;
      }
    }

    return (
      <div className={styles.documentCreator}>
        <h2 className={styles.heading}>{'\u05D9\u05E6\u05D9\u05E8\u05EA \u05DE\u05E1\u05DE\u05DA'}</h2>

        {userProfileError && (
          <MessageBar messageBarType={MessageBarType.warning} className={styles.messageBar}>
            {userProfileError}
          </MessageBar>
        )}

        {userProfile && (
          <div className={styles.userInfo}>
            <Label>{'\u05DE\u05D7\u05D5\u05D1\u05E8 \u05DB: '}{userProfile.displayName}</Label>
            {userProfile.department && <Label>{'\u05DE\u05D7\u05DC\u05E7\u05D4: '}{userProfile.department}</Label>}
          </div>
        )}

        <div className={styles.formSection}>
          <Dropdown
            label={'\u05E1\u05D5\u05D2 \u05DE\u05E1\u05DE\u05DA'}
            selectedKey={selectedDocType.key}
            options={this._docTypeOptions}
            onChange={this._onDocTypeChange}
            disabled={isCreating}
            onRenderOption={this._onRenderDocTypeOption}
            onRenderTitle={this._onRenderDocTypeTitle}
          />
        </div>

        <div className={styles.formSection}>
          <TextField
            label={'\u05E9\u05DD \u05DE\u05E1\u05DE\u05DA'}
            placeholder={'\u05D4\u05D6\u05DF \u05E9\u05DD \u05DE\u05E1\u05DE\u05DA (\u05DC\u05DC\u05D0 \u05E1\u05D9\u05D5\u05DE\u05EA)'}
            value={documentName}
            onChange={this._onDocumentNameChange}
            disabled={isCreating}
          />
          {fileNamePreview && (
            <div className={styles.filePreview}>
              <Icon iconName={selectedDocType.icon} className={styles.filePreviewIcon} />
              <span>{fileNamePreview}</span>
            </div>
          )}
        </div>

        <div className={styles.formSection}>
          <Label>{'\u05D1\u05D7\u05E8 \u05EA\u05D9\u05E7\u05D9\u05D9\u05EA \u05D9\u05E2\u05D3:'}</Label>
          {selectedNode && (
            <div className={styles.selectedLocation}>
              {'\u05E0\u05D1\u05D7\u05E8: '}<strong>{selectedNode.serverRelativeUrl || selectedNode.title}</strong>
            </div>
          )}
          <div className={styles.treeContainer}>
            <LocationTreePicker
              rootNode={rootNode}
              selectedNodeId={selectedNode ? selectedNode.id : undefined}
              onNodeSelect={this._onNodeSelect}
              onNodeExpand={this._onNodeExpand}
            />
          </div>
        </div>

        <div className={styles.formSection}>
          <PrimaryButton
            text={isCreating ? '\u05D9\u05D5\u05E6\u05E8...' : '\u05E6\u05D5\u05E8 \u05DE\u05E1\u05DE\u05DA'}
            onClick={this._onCreateDocument}
            disabled={!canCreate}
          />
          {isCreating && <Spinner size={SpinnerSize.small} className={styles.inlineSpinner} />}
        </div>

        {errorMessage && (
          <MessageBar
            messageBarType={MessageBarType.error}
            className={styles.messageBar}
            onDismiss={this._dismissError}
          >
            {errorMessage}
          </MessageBar>
        )}

        {successMessage && (
          <MessageBar
            messageBarType={MessageBarType.success}
            className={styles.messageBar}
            onDismiss={this._dismissSuccess}
          >
            {successMessage}
            {createdFileUrl && (
              <div>
                <Link href={createdFileUrl} target="_blank">{'\u05E4\u05EA\u05D7 \u05E7\u05D5\u05D1\u05E5'}</Link>
              </div>
            )}
          </MessageBar>
        )}
      </div>
    );
  }

  private _onRenderDocTypeOption = (option?: IDropdownOption): JSX.Element | null => {
    if (!option) return null;
    return (
      <div className={styles.docTypeOption}>
        <Icon iconName={option.data ? option.data.icon : 'Document'} className={styles.docTypeIcon} />
        <span>{option.text}</span>
      </div>
    );
  }

  private _onRenderDocTypeTitle = (options?: IDropdownOption[]): JSX.Element | null => {
    if (!options || options.length === 0) return null;
    const option = options[0];
    return (
      <div className={styles.docTypeOption}>
        <Icon iconName={option.data ? option.data.icon : 'Document'} className={styles.docTypeIcon} />
        <span>{option.text}</span>
      </div>
    );
  }

  private _onDocTypeChange = (_ev: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (!option) return;
    let found: IDocumentType | undefined;
    for (let i = 0; i < DocumentTypes.length; i++) {
      if (DocumentTypes[i].key === option.key) {
        found = DocumentTypes[i];
        break;
      }
    }
    if (found) {
      this.setState({ selectedDocType: found });
    }
  }

  private _onDocumentNameChange = (_ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
    this.setState({
      documentName: newValue || '',
      errorMessage: '',
      successMessage: '',
      createdFileUrl: ''
    });
  }

  private _onNodeSelect = (node: ITreeNode): void => {
    this.setState({
      selectedNode: node,
      errorMessage: '',
      successMessage: '',
      createdFileUrl: ''
    });
  }

  private _onNodeExpand = async (node: ITreeNode): Promise<void> => {
    this._updateNodeInTree(this.state.rootNode, node.id, { isLoading: true });
    this.forceUpdate();

    try {
      let children: ITreeNode[] = [];

      if (node.type === 'site') {
        children = await this.props.treeService.loadSiteChildren(node.webUrl);
      } else if (node.type === 'library') {
        children = await this.props.treeService.loadLibraryFolders(node.webUrl, node.serverRelativeUrl!);
      } else if (node.type === 'folder') {
        children = await this.props.treeService.loadFolderChildren(node.webUrl, node.serverRelativeUrl!);
      }

      this._updateNodeInTree(this.state.rootNode, node.id, {
        children: children,
        isLoaded: true,
        isLoading: false
      });
      this.forceUpdate();
    } catch (error) {
      console.error('Failed to expand node:', error);
      this._updateNodeInTree(this.state.rootNode, node.id, { isLoading: false });
      this.forceUpdate();
    }
  }

  private _updateNodeInTree(node: ITreeNode, targetId: string, updates: Partial<ITreeNode>): boolean {
    if (node.id === targetId) {
      const keys = Object.keys(updates);
      for (let i = 0; i < keys.length; i++) {
        const key = keys[i];
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (node as any)[key] = (updates as any)[key];
      }
      return true;
    }
    if (node.children) {
      for (let j = 0; j < node.children.length; j++) {
        if (this._updateNodeInTree(node.children[j], targetId, updates)) {
          return true;
        }
      }
    }
    return false;
  }

  private _dismissError = (): void => {
    this.setState({ errorMessage: '' });
  }

  private _dismissSuccess = (): void => {
    this.setState({ successMessage: '', createdFileUrl: '' });
  }

  private _onCreateDocument = async (): Promise<void> => {
    const { selectedNode, documentName, selectedDocType, userProfile } = this.state;

    if (!selectedNode || !selectedNode.serverRelativeUrl || !userProfile) {
      return;
    }

    this.setState({ isCreating: true, errorMessage: '', successMessage: '', createdFileUrl: '' });

    try {
      const fileRelativeUrl = await this.props.documentService.createDocument(
        selectedNode.webUrl,
        selectedNode.serverRelativeUrl,
        documentName,
        selectedDocType,
        userProfile
      );

      // Build a link to the file - use Office Online URL for Office documents
      let origin = selectedNode.webUrl;
      try {
        origin = new URL(selectedNode.webUrl).origin;
      } catch (ex) {
        console.debug('URL parse fallback', ex);
      }

      let fileAbsoluteUrl: string;
      if (selectedDocType.onlinePrefix) {
        // Office Online format: {origin}/:w:/r{serverRelativeUrl}?csf=1&web=1
        fileAbsoluteUrl = origin + '/' + selectedDocType.onlinePrefix + '/r' + fileRelativeUrl + '?csf=1&web=1';
      } else {
        fileAbsoluteUrl = origin + fileRelativeUrl;
      }

      this.setState({
        isCreating: false,
        successMessage: '\u05D4\u05DE\u05E1\u05DE\u05DA \u05E0\u05D5\u05E6\u05E8 \u05D1\u05D4\u05E6\u05DC\u05D7\u05D4!',
        createdFileUrl: fileAbsoluteUrl,
        documentName: ''
      });
    } catch (error) {
      this.setState({
        isCreating: false,
        errorMessage: error instanceof Error ? error.message : String(error)
      });
    }
  }
}
