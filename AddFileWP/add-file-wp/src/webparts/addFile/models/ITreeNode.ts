export interface ITreeNode {
  id: string;
  title: string;
  type: 'site' | 'library' | 'folder';
  webUrl: string;
  serverRelativeUrl?: string;
  children?: ITreeNode[];
  isLoaded?: boolean;
  isLoading?: boolean;
  selectable: boolean;
}
