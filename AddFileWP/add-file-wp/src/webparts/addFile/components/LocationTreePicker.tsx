import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { ITreeNode } from '../models/ITreeNode';
import styles from './AddFile.module.scss';

export interface ILocationTreePickerProps {
  rootNode: ITreeNode;
  selectedNodeId: string | undefined;
  onNodeSelect: (node: ITreeNode) => void;
  onNodeExpand: (node: ITreeNode) => void;
}

interface ITreeNodeItemProps {
  node: ITreeNode;
  depth: number;
  selectedNodeId: string | undefined;
  onNodeSelect: (node: ITreeNode) => void;
  onNodeExpand: (node: ITreeNode) => void;
}

const TreeNodeItem: React.FC<ITreeNodeItemProps> = (props) => {
  const { node, depth, selectedNodeId, onNodeSelect, onNodeExpand } = props;
  const [isExpanded, setIsExpanded] = React.useState(false);

  const getIcon = (): string => {
    switch (node.type) {
      case 'site': return 'Globe';
      case 'library': return 'DocLibrary';
      case 'folder': return isExpanded ? 'FabricOpenFolderHorizontal' : 'FabricFolder';
      default: return 'Folder';
    }
  };

  const handleToggle = (): void => {
    if (!isExpanded && !node.isLoaded) {
      onNodeExpand(node);
    }
    setIsExpanded(!isExpanded);
  };

  const handleSelect = (): void => {
    if (node.selectable) {
      onNodeSelect(node);
    }
  };

  const isSelected = selectedNodeId === node.id;

  return (
    <div>
      <div
        className={`${styles.treeNode} ${isSelected ? styles.treeNodeSelected : ''} ${node.selectable ? styles.treeNodeSelectable : ''}`}
        style={{ paddingLeft: `${depth * 20 + 4}px` }}
      >
        <span className={styles.treeToggle} onClick={handleToggle}>
          {node.isLoading ? (
            <Spinner size={SpinnerSize.xSmall} />
          ) : (
            <Icon iconName={isExpanded ? 'ChevronDown' : 'ChevronRight'} className={styles.treeChevron} />
          )}
        </span>
        <span className={styles.treeLabel} onClick={handleSelect} onDoubleClick={handleToggle}>
          <Icon iconName={getIcon()} className={styles.treeIcon} />
          <span className={styles.treeTitle}>{node.title}</span>
        </span>
      </div>
      {isExpanded && node.children && node.children.length > 0 && (
        <div>
          {node.children.map((child) => (
            <TreeNodeItem
              key={child.id}
              node={child}
              depth={depth + 1}
              selectedNodeId={selectedNodeId}
              onNodeSelect={onNodeSelect}
              onNodeExpand={onNodeExpand}
            />
          ))}
        </div>
      )}
    </div>
  );
};

export const LocationTreePicker: React.FC<ILocationTreePickerProps> = (props) => {
  const { rootNode, selectedNodeId, onNodeSelect, onNodeExpand } = props;

  return (
    <div className={styles.treePicker}>
      <TreeNodeItem
        node={rootNode}
        depth={0}
        selectedNodeId={selectedNodeId}
        onNodeSelect={onNodeSelect}
        onNodeExpand={onNodeExpand}
      />
    </div>
  );
};
