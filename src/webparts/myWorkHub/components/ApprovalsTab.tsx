import * as React from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { GraphBetaCall } from './IMyWorkHubProps';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Dialog, DialogFooter } from '@fluentui/react/lib/Dialog';
import { TextField } from '@fluentui/react/lib/TextField';
import ReactMarkdown from 'react-markdown';

export interface IApprovalsTabProps {
  msGraphClient?: MSGraphClientV3;
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
  owner?: { user?: { id?: string; displayName?: string } };
}

interface OwnerDetails {
  displayName: string;
  photoUrl: string | null;
}

interface RejectDialogState {
  open: boolean;
  approvalId: string | null;
  reason: string;
  submitting: boolean;
}

interface IApprovalsTabState {
  items: IApprovalItem[];
  loading: boolean;
  error?: string;
  ownerDetails: Record<string, OwnerDetails>;
  rejectDialog: RejectDialogState;
}

export const ApprovalsTab: React.FC<IApprovalsTabProps> = (props) => {
  const { msGraphClient, callGraphBeta, onError } = props;
  const [state, setState] = React.useState<IApprovalsTabState>({
    items: [],
    loading: true,
    ownerDetails: {},
    rejectDialog: { open: false, approvalId: null, reason: '', submitting: false }
  });
  const photoObjectUrlsRef = React.useRef<string[]>([]);

  const loadOwnerDetails = React.useCallback(
    async (userId: string): Promise<OwnerDetails> => {
      if (!msGraphClient) return { displayName: '', photoUrl: null };
      try {
        const user = (await msGraphClient
          .api(`/users/${userId}?$select=displayName`)
          .get()) as { displayName?: string };
        let photoUrl: string | null = null;
        try {
          const api = (msGraphClient as { api: (path: string) => { responseType: (t: string) => { get: () => Promise<Blob> } } }).api;
          const blob = await api(`/users/${userId}/photo/$value`).responseType('blob').get();
          if (blob && blob.size > 0) {
            const url = URL.createObjectURL(blob);
            photoObjectUrlsRef.current.push(url);
            photoUrl = url;
          }
        } catch {
          // User may have no photo
        }
        return { displayName: user.displayName ?? '', photoUrl };
      } catch {
        return { displayName: '', photoUrl: null };
      }
    },
    [msGraphClient]
  );

  const loadApprovals = React.useCallback(async () => {
    if (!callGraphBeta) {
      setState(prev => ({ ...prev, loading: false, error: 'Graph beta not available.' }));
      return;
    }
    setState(prev => ({ ...prev, loading: true, error: undefined }));
    try {
      const result = await callGraphBeta("/solutions/approval/approvalItems?$orderby=createdDateTime desc&$top=6");
      const raw = (result.value || []) as IApprovalItem[];
      const value = raw.filter((a) => a.state === 'pending');
      setState(prev => ({ ...prev, items: value, loading: false, ownerDetails: {} }));

      const ownerIds: string[] = Array.from(new Set(value.map((a) => a.owner?.user?.id).filter((id): id is string => !!id)));
      if (ownerIds.length === 0 || !msGraphClient) return;
      const details: Record<string, OwnerDetails> = {};
      await Promise.all(
        ownerIds.map(async (id) => {
          details[id] = await loadOwnerDetails(id);
        })
      );
      setState(prev => ({ ...prev, ownerDetails: { ...prev.ownerDetails, ...details } }));
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(`Approvals: ${message}`);
      setState(prev => ({ ...prev, loading: false, error: message }));
    }
  }, [callGraphBeta, msGraphClient, loadOwnerDetails, onError]);

  React.useEffect(() => {
    return () => {
      photoObjectUrlsRef.current.forEach((url) => URL.revokeObjectURL(url));
      photoObjectUrlsRef.current = [];
    };
  }, []);

  React.useEffect(() => {
    void loadApprovals();
  }, [loadApprovals]);

  const respond = async (approvalId: string, response: string, comments?: string): Promise<void> => {
    if (!callGraphBeta) return;
    try {
      await callGraphBeta(`/solutions/approval/approvalItems/${approvalId}/responses`, {
        method: 'POST',
        body: { response, comments: comments ?? '' }
      });
      setState(prev => ({
        ...prev,
        items: prev.items.filter(i => i.id !== approvalId),
        rejectDialog: { ...prev.rejectDialog, open: false, approvalId: null, reason: '', submitting: false }
      }));
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(`Respond: ${message}`);
      setState(prev => ({ ...prev, rejectDialog: { ...prev.rejectDialog, submitting: false } }));
    }
  };

  const openRejectDialog = (approvalId: string): void => {
    setState(prev => ({
      ...prev,
      rejectDialog: { open: true, approvalId, reason: '', submitting: false }
    }));
  };

  const closeRejectDialog = (): void => {
    setState(prev => ({
      ...prev,
      rejectDialog: { open: false, approvalId: null, reason: '', submitting: false }
    }));
  };

  const submitReject = (): void => {
    const { approvalId, reason } = state.rejectDialog;
    if (!approvalId) return;
    setState(prev => ({ ...prev, rejectDialog: { ...prev.rejectDialog, submitting: true } }));
    void respond(approvalId, 'Reject', reason.trim() || undefined);
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

  const { rejectDialog } = state;

  return (
    <>
      <Dialog
        hidden={!rejectDialog.open}
        onDismiss={closeRejectDialog}
        dialogContentProps={{
          title: 'Reject approval',
          subText: 'Optionally provide a reason for the requester.'
        }}
        modalProps={{ isBlocking: true }}
      >
        <TextField
          label="Reason (optional)"
          multiline
          rows={3}
          value={rejectDialog.reason}
          onChange={(_ev, newValue) =>
            setState(prev => ({
              ...prev,
              rejectDialog: { ...prev.rejectDialog, reason: newValue ?? '' }
            }))
          }
          placeholder="Enter a reason for rejecting..."
          disabled={rejectDialog.submitting}
        />
        <DialogFooter>
          <PrimaryButton text="Reject" onClick={submitReject} disabled={rejectDialog.submitting} />
          <DefaultButton text="Cancel" onClick={closeRejectDialog} disabled={rejectDialog.submitting} />
        </DialogFooter>
      </Dialog>
      <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
      {state.items.map(item => {
        const ownerId = item.owner?.user?.id;
        const owner = ownerId ? state.ownerDetails[ownerId] : undefined;
        const ownerName = owner?.displayName || item.owner?.user?.displayName;
        return (
          <li key={item.id} style={{ marginBottom: '16px', padding: '12px', border: '1px solid #edebe9', borderRadius: '4px' }}>
            <div style={{ fontWeight: 600 }}>{item.displayName}</div>
            {item.description && (
              <div
                className="approval-description"
                style={{ fontSize: '13px', color: '#323130', marginTop: '6px', lineHeight: 1.5 }}
              >
                <ReactMarkdown>{item.description}</ReactMarkdown>
              </div>
            )}
            {(ownerName || owner?.photoUrl) && (
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', marginTop: '8px', fontSize: '12px', color: '#605e5c' }}>
                {owner?.photoUrl ? (
                  <img
                    src={owner.photoUrl}
                    alt=""
                    style={{ width: 24, height: 24, borderRadius: '50%', objectFit: 'cover' }}
                  />
                ) : (
                  <span style={{ width: 24, height: 24, borderRadius: '50%', background: '#edebe9', display: 'inline-block' }} />
                )}
                {ownerName && <span>From: {ownerName}</span>}
              </div>
            )}
            {item.createdDateTime && (
              <div style={{ fontSize: '12px', marginTop: '4px' }}>
                {new Date(item.createdDateTime).toLocaleString()}
              </div>
            )}
            <div style={{ marginTop: '8px', display: 'flex', gap: '8px' }}>
              {(item.responseOptions || ['Approve', 'Reject']).map(opt =>
                opt === 'Approve' ? (
                  <PrimaryButton key={opt} text={opt} onClick={() => respond(item.id, opt)} />
                ) : (
                  <DefaultButton key={opt} text={opt} onClick={() => openRejectDialog(item.id)} />
                )
              )}
            </div>
          </li>
        );
      })}
    </ul>
    </>
  );
};
