import * as React from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { GraphBetaCall } from './IMyWorkHubProps';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

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
      const contextParts: string[] = [];
      try {
        const todoLists = await msGraphClient.api('/me/todo/lists').get() as { value?: Array<{ id: string; displayName: string }> };
        let taskCount = 0;
        const taskTitles: string[] = [];
        for (const list of todoLists.value || []) {
          const tasksRes = await msGraphClient.api(`/me/todo/lists/${list.id}/tasks`).get() as { value?: Array<{ title: string; status: string }> };
          for (const t of tasksRes.value || []) {
            if (t.status !== 'completed') {
              taskCount++;
              if (taskTitles.length < 5) taskTitles.push(t.title);
            }
          }
        }
        const plannerRes = await msGraphClient.api('/me/planner/tasks').get() as { value?: Array<{ title: string; percentComplete?: number }> };
        for (const t of plannerRes.value || []) {
          if (t.percentComplete !== 100) {
            taskCount++;
            if (taskTitles.length < 5) taskTitles.push(t.title);
          }
        }
        contextParts.push(`Tasks: ${taskCount} pending (e.g. ${taskTitles.slice(0, 3).join(', ') || 'none'}).`);
      } catch {
        contextParts.push('Tasks: could not load.');
      }
      try {
        // Graph doesn't allow $filter on approvals; order by createdDateTime and filter client-side
        const approvals = await callGraphBeta("/solutions/approval/approvalItems?$orderby=createdDateTime desc&$top=50");
        const raw = (approvals.value || []) as Array<{ displayName?: string; state?: string }>;
        const pending = raw.filter((a) => a.state === 'pending');
        const count = pending.length;
        const names = pending.slice(0, 3).map((a) => a.displayName);
        contextParts.push(`Pending approvals: ${count} (e.g. ${names.join(', ') || 'none'}).`);
      } catch {
        contextParts.push('Approvals: could not load.');
      }
      try {
        const recent = await msGraphClient.api('/me/drive/recent?$top=5').get() as { value?: Array<{ name?: string; resourceVisualization?: { title?: string } }> };
        const titles = (recent.value || []).map(v => v.name ?? v.resourceVisualization?.title).filter(Boolean);
        contextParts.push(`Recent files: ${titles.length} (e.g. ${titles.slice(0, 2).join(', ') || 'none'}).`);
      } catch {
        contextParts.push('Recent files: could not load.');
      }
      const contextString = `The user has: ${contextParts.join(' ')}`;

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
          message: { text: 'Summarize my pending work today.' },
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
      <p>Get a short summary of your pending work using Microsoft 365 Copilot.</p>
      <PrimaryButton text="Summarize my pending work today" onClick={runSummarize} disabled={loading} />
      {loading && <Spinner style={{ marginTop: 8 }} />}
      {error && (
        <MessageBar messageBarType={MessageBarType.warning} style={{ marginTop: 8 }}>
          {error} (User needs Microsoft 365 Copilot license and API permissions.)
        </MessageBar>
      )}
      {summary && (
        <div style={{ marginTop: 16, padding: 12, backgroundColor: '#f3f2f1', borderRadius: 4, whiteSpace: 'pre-wrap' }}>
          {summary}
        </div>
      )}
    </div>
  );
};
