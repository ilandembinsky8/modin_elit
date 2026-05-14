import * as React from 'react';
import { TextField } from '@fluentui/react/lib/TextField';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Label } from '@fluentui/react/lib/Label';
import { Link } from '@fluentui/react/lib/Link';
import { LocationTreePicker } from './LocationTreePicker';
import { ITreeNode } from '../models/ITreeNode';
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
  isCreating: boolean;
  errorMessage: string;
  successMessage: string;
  createdFileUrl: string;
  userProfile: IUserProfile | undefined;
  userProfileError: string;
}

export class DocumentCreator extends React.Component<IDocumentCreatorProps, IDocumentCreatorState> {
  constructor(props: IDocumentCreatorProps) {
    super(props);

    const rootNode = props.treeService.getRootNode(props.rootSiteUrl, props.rootSiteTitle);

    this.state = {
      rootNode,
      selectedNode: undefined,
      documentName: '',
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
    } catch (error) {
      this.setState({
        userProfileError: `Failed to load user profile: ${error instanceof Error ? error.message : String(error)}`
      });
    }
  }

  public render(): React.ReactElement {
    const {
      rootNode,
      selectedNode,
      documentName,
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

    return (
      <div className={styles.documentCreator}>
        <h2 className={styles.heading}>Create Document</h2>

        {userProfileError && (
          <MessageBar messageBarType={MessageBarType.warning} className={styles.messageBar}>
            {userProfileError}
          </MessageBar>
        )}

        {userProfile && (
          <div className={styles.userInfo}>
            <Label>Logged in as: {userProfile.displayName}</Label>
            {userProfile.department && <Label>Department: {userProfile.department}</Label>}
          </div>
        )}

        <div className={styles.formSection}>
          <TextField
            label="Document Name"
            placeholder="Enter document name (e.g., MyDocument or MyDocument.txt)"
            value={documentName}
            onChange={this._onDocumentNameChange}
            disabled={isCreating}
          />
        </div>

        <div className={styles.formSection}>
          <Label>Select target folder:</Label>
          {selectedNode && (
            <div className={styles.selectedLocation}>
              Selected: <strong>{selectedNode.serverRelativeUrl || selectedNode.title}</strong>
            </div>
          )}
          <div className={styles.treeContainer}>
            <LocationTreePicker
              rootNode={rootNode}
              selectedNodeId={selectedNode?.id}
              onNodeSelect={this._onNodeSelect}
              onNodeExpand={this._onNodeExpand}
            />
          </div>
        </div>

        <div className={styles.formSection}>
          <PrimaryButton
            text={isCreating ? 'Creating...' : 'Create Document'}
            onClick={this._onCreateDocument}
            disabled={!canCreate}
          />
          {isCreating && <Spinner size={SpinnerSize.small} className={styles.inlineSpinner} />}
        </div>

        {errorMessage && (
          <MessageBar
            messageBarType={MessageBarType.error}
            className={styles.messageBar}
            onDismiss={() => this.setState({ errorMessage: '' })}
          >
            {errorMessage}
          </MessageBar>
        )}

        {successMessage && (
          <MessageBar
            messageBarType={MessageBarType.success}
            className={styles.messageBar}
            onDismiss={() => this.setState({ successMessage: '', createdFileUrl: '' })}
          >
            {successMessage}
            {createdFileUrl && (
              <div>
                <Link href={createdFileUrl} target="_blank">Open file</Link>
              </div>
            )}
          </MessageBar>
        )}
      </div>
    );
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
        children,
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
      const keys = Object.keys(updates) as Array<keyof ITreeNode>;
      for (const key of keys) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        (node as any)[key] = (updates as any)[key];
      }
      return true;
    }
    if (node.children) {
      for (const child of node.children) {
        if (this._updateNodeInTree(child, targetId, updates)) {
          return true;
        }
      }
    }
    return false;
  }

  private _onCreateDocument = async (): Promise<void> => {
    const { selectedNode, documentName, userProfile } = this.state;

    if (!selectedNode || !selectedNode.serverRelativeUrl || !userProfile) {
      return;
    }

    this.setState({ isCreating: true, errorMessage: '', successMessage: '', createdFileUrl: '' });

    try {
      const fileRelativeUrl = await this.props.documentService.createDocument(
        selectedNode.webUrl,
        selectedNode.serverRelativeUrl,
        documentName,
        userProfile
      );

      // Build a link to the file
      const origin = new URL(selectedNode.webUrl).origin;
      const fileAbsoluteUrl = `${origin}${fileRelativeUrl}`;

      this.setState({
        isCreating: false,
        successMessage: `Document created successfully!`,
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
