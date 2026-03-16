import * as React from 'react';
import type { GraphBetaCall } from './IMyWorkHubProps';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

export interface IApprovalsTabProps {
  callGraphBeta?: GraphBetaCall;
  onError?: (message: string) => void;
}

interface IApprovalItem {
  id: string;
  displayName: string;
  description?: string;
  createdDateTime?: string;
  state?: string;
  responseOptions?: string[];
  owner?: { user?: { displayName?: string } };
}

interface IApprovalsTabState {
  items: IApprovalItem[];
  loading: boolean;
  error?: string;
}

export const ApprovalsTab: React.FunctionComponent<IApprovalsTabProps> = (props) => {
  const { callGraphBeta, onError } = props;
  const [state, setState] = React.useState<IApprovalsTabState>({ items: [], loading: true });

  const loadApprovals = React.useCallback(async () => {
    if (!callGraphBeta) {
      setState(prev => ({ ...prev, loading: false, error: 'Graph beta not available.' }));
      return;
    }
    setState(prev => ({ ...prev, loading: true, error: undefined }));
    try {
      const result = await callGraphBeta("/solutions/approval/approvalItems?$filter=state eq 'pending'&$top=50");
      const value = (result.value || []) as IApprovalItem[];
      setState({ items: value, loading: false });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(`Approvals: ${message}`);
      setState(prev => ({ ...prev, loading: false, error: message }));
    }
  }, [callGraphBeta, onError]);

  React.useEffect(() => {
    void loadApprovals();
  }, [loadApprovals]);

  const respond = async (approvalId: string, response: string): Promise<void> => {
    if (!callGraphBeta) return;
    try {
      await callGraphBeta(`/solutions/approval/approvalItems/${approvalId}/responses`, {
        method: 'POST',
        body: { response, comments: '' }
      });
      setState(prev => ({
        ...prev,
        items: prev.items.filter(i => i.id !== approvalId)
      }));
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(`Respond: ${message}`);
    }
  };

  if (state.loading) {
    return <Spinner label="Loading approvals..." />;
  }

  if (state.error) {
    return (
      <MessageBar messageBarType={MessageBarType.warning}>
        {state.error} (Check permissions or tenant support for Approvals app.)
      </MessageBar>
    );
  }

  if (state.items.length === 0) {
    return <p>No pending approvals.</p>;
  }

  return (
    <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
      {state.items.map(item => (
        <li key={item.id} style={{ marginBottom: '16px', padding: '12px', border: '1px solid #edebe9', borderRadius: '4px' }}>
          <div style={{ fontWeight: 600 }}>{item.displayName}</div>
          {item.description && (
            <div style={{ fontSize: '12px', color: '#605e5c', marginTop: '4px' }}>
              {item.description.length > 200 ? `${item.description.substring(0, 200)}...` : item.description}
            </div>
          )}
          {item.owner?.user?.displayName && (
            <div style={{ fontSize: '12px', marginTop: '4px' }}>From: {item.owner.user.displayName}</div>
          )}
          {item.createdDateTime && (
            <div style={{ fontSize: '12px', marginTop: '4px' }}>
              {new Date(item.createdDateTime).toLocaleString()}
            </div>
          )}
          <div style={{ marginTop: '8px', display: 'flex', gap: '8px' }}>
            {(item.responseOptions || ['Approve', 'Reject']).map(opt => (
              opt === 'Approve' ? (
                <PrimaryButton key={opt} text={opt} onClick={() => respond(item.id, opt)} />
              ) : (
                <DefaultButton key={opt} text={opt} onClick={() => respond(item.id, opt)} />
              )
            ))}
          </div>
        </li>
      ))}
    </ul>
  );
};
