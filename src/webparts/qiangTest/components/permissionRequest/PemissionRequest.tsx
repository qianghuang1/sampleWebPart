import * as React from 'react';
import { Icon, PrimaryButton } from '@microsoft/office-ui-fabric-react-bundle';

import styles from './PermissionRequest.module.scss';

export interface IPermissionRequestProps {
  tenantId: string;
  clientId: string;
}

export default class PermissionRequest extends React.Component<IPermissionRequestProps, {}> {
  public render(): JSX.Element {
    return <div className={styles.mainContainer}>
      <div className={styles.grantPermissionWarningContent}>
        <Icon iconName={'Warning'} className={styles.warningIcon} />
        <div>You haven't grant permission for this low trust app, please click the button below to grant your consent.</div>
      </div>
      <div className={styles.consentButton}>
        <PrimaryButton
          onClick={() => {
            window.open(`https://login.microsoftonline.com/${this.props.tenantId}/oauth2/v2.0/authorize?response_type=id_token%20token&scope=https://graph.microsoft.com/Sites.Read.All openid profile&client_id=${this.props.clientId}&&redirect_uri=http://localhost:3000/success`, '_blank');
          }}
          text={'Consent in the AAD'} />
      </div>
    </div>;
  }
}