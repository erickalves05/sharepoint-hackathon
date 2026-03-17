import * as React from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { GraphBetaCall } from './IMyWorkHubProps';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import ReactMarkdown from 'react-markdown';

export interface ISummaryTabProps {
  msGraphClient: MSGraphClientV3;
  callGraphBeta?: GraphBetaCall;
  getSummaryContext: () => Promise<string>;
  onError?: (message: string) => void;
}

interface ICopilotMessage {
  text?: string;
}

interface ICopilotChatResponse {
  messages?: ICopilotMessage[];
}

export const SummaryTab: React.FC<ISummaryTabProps> = (props) => {
  const { msGraphClient, callGraphBeta, onError } = props;
  const [loading, setLoading] = React.useState(false);
  const [summary, setSummary] = React.useState<string | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);

  const runSummarize = async (): Promise<void> => {
    if (!callGraphBeta) {
      setError('Copilot (Graph beta) not available.');
      return;
    }
    setLoading(true);
    setError(undefined);
    setSummary(undefined);
    try {
      const taskLines: string[] = [];
      try {
        type TodoTask = { title: string; status: string; dueDateTime?: { dateTime: string }; linkedResources?: Array<{ webUrl?: string }> };
        const todoLists = await msGraphClient.api('/me/todo/lists').get() as { value?: Array<{ id: string; displayName: string }> };
        for (const list of todoLists.value || []) {
          const tasksRes = await msGraphClient
            .api(`/me/todo/lists/${list.id}/tasks?$expand=linkedResources`)
            .get() as { value?: TodoTask[] };
          for (const t of tasksRes.value || []) {
            if (t.status !== 'completed') {
              const source = t.linkedResources?.[0]?.webUrl ? 'Flagged mail' : 'To Do';
              const due = t.dueDateTime?.dateTime ? `, due ${new Date(t.dueDateTime.dateTime).toLocaleDateString()}` : '';
              const title = t.title.replace(/\s*\(.*$/, '').trim();
              taskLines.push(`- "${title}" (${source}, list: ${list.displayName}${due})`);
            }
          }
        }
        type PlannerTask = { title: string; percentComplete?: number; dueDateTime?: string };
        const plannerRes = await msGraphClient.api('/me/planner/tasks').get() as { value?: PlannerTask[] };
        for (const t of plannerRes.value || []) {
          if (t.percentComplete !== 100) {
            const due = t.dueDateTime ? `, due ${new Date(t.dueDateTime).toLocaleDateString()}` : '';
            taskLines.push(`- "${t.title}" (Planner${due})`);
          }
        }
        if (taskLines.length > 0) {
          taskLines.unshift('TASKS:', '');
        } else {
          taskLines.push('TASKS:', 'None pending.');
        }
      } catch {
        taskLines.push('TASKS:', 'Could not load.');
      }

      const approvalLines: string[] = [];
      try {
        const approvals = await callGraphBeta("/solutions/approval/approvalItems?$orderby=createdDateTime desc&$top=50");
        type ApprovalItem = { displayName?: string; state?: string; createdDateTime?: string; owner?: { user?: { displayName?: string } } };
        const raw = (approvals.value || []) as ApprovalItem[];
        const pending = raw.filter((a) => a.state === 'pending');
        approvalLines.push('PENDING APPROVALS:', '');
        if (pending.length > 0) {
          for (const a of pending) {
            const owner = "Erick Alves" //a.owner?.user?.displayName ?? 'Unknown';
            const created = a.createdDateTime ? new Date(a.createdDateTime).toLocaleDateString() : 'unknown date';
            approvalLines.push(`- "${a.displayName ?? 'Untitled'}" (requested by ${owner}, created ${created})`);
          }
        } else {
          approvalLines.push('None pending.');
        }
      } catch {
        approvalLines.push('PENDING APPROVALS:', 'Could not load.');
      }

      const recentLines: string[] = [];
      try {
        type DriveItem = { name?: string; lastModifiedDateTime?: string };
        const recent = await msGraphClient.api('/me/drive/recent?$top=10').get() as { value?: DriveItem[] };
        const items = (recent.value || []).filter((v) => v.name);
        recentLines.push('RECENT FILES:', '');
        if (items.length > 0) {
          for (const v of items) {
            const modified = v.lastModifiedDateTime ? `, modified ${new Date(v.lastModifiedDateTime).toLocaleDateString()}` : '';
            recentLines.push(`- ${v.name}${modified}`);
          }
        } else {
          recentLines.push('None.');
        }
      } catch {
        recentLines.push('RECENT FILES:', 'Could not load.');
      }

      const contextString = [
        'PENDING WORK DATA (prioritize this data). Feel free to use emojis and other formatting to make the summary more engaging:',
        '',
        ...taskLines,
        '',
        ...approvalLines,
        '',
        ...recentLines
      ].join('\n');

      const createRes = await callGraphBeta('/copilot/conversations', { method: 'POST', body: {} }) as { id?: string };
      const conversationId = createRes.id;
      if (!conversationId) {
        setError('Could not create Copilot conversation. Check Copilot license and permissions.');
        setLoading(false);
        return;
      }

      const chatRes = await callGraphBeta(`/copilot/conversations/${conversationId}/chat`, {
        method: 'POST',
        body: {
          message: {
            text: 'Summarize my pending work today. Prioritize the data in the additional context (tasks, approvals, recent files).'
          },
          additionalContext: [{ text: contextString }],
          locationHint: { timeZone: 'UTC' }
        }
      }) as ICopilotChatResponse;

      const messages = chatRes.messages || [];
      const lastMessage = messages.length > 0 ? messages[messages.length - 1] : undefined;
      const text = lastMessage?.text;
      if (text) {
        setSummary(text);
      } else {
        setError('No response from Copilot. Check permissions and license.');
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(`Summary: ${message}`);
      setError(message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div>
      {!summary && !loading && (
        <div
          style={{
            padding: '24px',
            border: '1px solid #edebe9',
            borderRadius: '4px',
            backgroundColor: '#faf9f8',
            marginBottom: '16px'
          }}
        >
          <div style={{ display: 'flex', alignItems: 'flex-start', gap: '16px' }}>
            <Icon
              iconName="Lightbulb"
              styles={{ root: { fontSize: 36, flexShrink: 0, marginTop: '4px' } }}
            />
            <div>
              <div style={{ fontSize: '18px', fontWeight: 600, marginBottom: '8px' }}>
                Summarize your pending work
              </div>
              <p style={{ margin: 0, color: '#605e5c', lineHeight: 1.5, marginBottom: '16px' }}>
                Get a short summary of your tasks, approvals, and recent files using Microsoft 365 Copilot.
              </p>
              <PrimaryButton text="Summarize my pending work today" onClick={runSummarize} disabled={loading} />
            </div>
          </div>
        </div>
      )}
      {loading && <Spinner label="Summarizing..." style={{ marginTop: 8 }} />}
      {error && (
        <MessageBar messageBarType={MessageBarType.warning} style={{ marginTop: 8 }}>
          {error} (User needs Microsoft 365 Copilot license and API permissions.)
        </MessageBar>
      )}
      {summary && (
        <div
          style={{
            marginTop: 16,
            padding: 16,
            border: '1px solid #edebe9',
            borderRadius: 4,
            backgroundColor: '#fff',
            fontSize: 14,
            lineHeight: 1.6
          }}
        >
          <div style={{ fontWeight: 600, marginBottom: 12, color: '#323130' }}>Summary</div>
          <div className="summary-content">
            <ReactMarkdown>{summary}</ReactMarkdown>
          </div>
        </div>
      )}
    </div>
  );
};
