import * as React from 'react';
import styles from './GraphInfoComponent.module.scss';

export interface IGraphInfoComponentProps {
  displayName?: string;
  description?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
}

export default class GraphInfoComponent extends React.Component<IGraphInfoComponentProps, {}>{
  public render(): JSX.Element {
    return <div className={styles.mainContainer}>
      <div className={styles.siteDisplayName}>{this.props.displayName}</div>
      <div className={styles.siteDescription}>{this.props.description ? this.props.description : 'No description for this site'}</div>
      <div className={styles.siteCreatedDate}>{`Created on: ${this.props.createdDateTime}`}</div>
      <div className={styles.siteModifiedDate}>{`Last Modified on: ${this.props.lastModifiedDateTime}`}</div>
    </div>;
  }
}