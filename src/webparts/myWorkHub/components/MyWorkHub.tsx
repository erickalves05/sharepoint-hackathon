import * as React from 'react';
import styles from './MyWorkHub.module.scss';
import type { IMyWorkHubProps } from './IMyWorkHubProps';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { TasksTab } from './TasksTab';
import { ApprovalsTab } from './ApprovalsTab';
import { RecentTab } from './RecentTab';
import { SummaryTab } from './SummaryTab';

export default class MyWorkHub extends React.Component<IMyWorkHubProps> {
  public render(): React.ReactElement<IMyWorkHubProps> {
    const {
      title,
      hasTeamsContext,
      msGraphClient,
      callGraphBeta,
      errorMessage,
      onError,
      dismissError
    } = this.props;

    const notReady = !msGraphClient;

    return (
      <section className={`${styles.myWorkHub} ${hasTeamsContext ? styles.teams : ''}`}>
        {title && (
          <div className={styles.webPartTitle}>{title}</div>
        )}
        {errorMessage && (
          <MessageBar
            messageBarType={MessageBarType.error}
            onDismiss={dismissError}
            dismissButtonAriaLabel="Close"
          >
            {errorMessage}
          </MessageBar>
        )}

        {notReady ? (
          <div className={styles.loading}>
            <Spinner label="Loading My Work Hub..." />
          </div>
        ) : (
          <Pivot aria-label="My Work Hub tabs">
            <PivotItem headerText="Tasks" itemKey="tasks">
              <div className={styles.tabContent}>
                <TasksTab
                  msGraphClient={msGraphClient}
                  onError={onError}
                />
              </div>
            </PivotItem>
            <PivotItem headerText="Approvals" itemKey="approvals">
              <div className={styles.tabContent}>
                <ApprovalsTab
                  msGraphClient={msGraphClient}
                  callGraphBeta={callGraphBeta}
                  onError={onError}
                />
              </div>
            </PivotItem>
            <PivotItem headerText="Recent" itemKey="recent">
              <div className={styles.tabContent}>
                <RecentTab
                  msGraphClient={msGraphClient}
                  onError={onError}
                />
              </div>
            </PivotItem>
            <PivotItem headerText="Summary" itemKey="summary">
              <div className={styles.tabContent}>
                <SummaryTab
                  msGraphClient={msGraphClient}
                  callGraphBeta={callGraphBeta}
                  getSummaryContext={this._getSummaryContext}
                  onError={onError}
                />
              </div>
            </PivotItem>
          </Pivot>
        )}
      </section>
    );
  }

  /** Build context string from Tasks, Approvals, Recent for Copilot. SummaryTab will call this. */
  private _getSummaryContext: () => Promise<string> = async () => {
    return Promise.resolve('No data loaded yet. Open Tasks, Approvals, and Recent tabs first to build context.');
  };
}
