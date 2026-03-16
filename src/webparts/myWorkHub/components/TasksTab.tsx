import * as React from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import type { IUnifiedTask } from './tasks/ITaskItem';

export interface ITasksTabProps {
  msGraphClient: MSGraphClientV3;
  onError?: (message: string) => void;
}

interface ITasksTabState {
  tasks: IUnifiedTask[];
  loading: boolean;
}

export const TasksTab: React.FunctionComponent<ITasksTabProps> = (props) => {
  const [state, setState] = React.useState<ITasksTabState>({ tasks: [], loading: true });
  const { msGraphClient, onError } = props;

  const loadTasks = React.useCallback(async () => {
    setState(prev => ({ ...prev, loading: true }));
    try {
      const all: IUnifiedTask[] = [];

      const todoLists = await msGraphClient
        .api('/me/todo/lists')
        .get() as { value: Array<{ id: string; displayName: string }> };

      for (const list of todoLists.value || []) {
        const tasksRes = await msGraphClient
          .api(`/me/todo/lists/${list.id}/tasks`)
          .get() as { value: Array<{ id: string; title: string; status: string; dueDateTime?: { dateTime: string }; webUrl?: string }> };
        for (const t of tasksRes.value || []) {
          if (t.status !== 'completed') {
            all.push({
              id: t.id,
              taskId: t.id,
              listId: list.id,
              title: t.title,
              dueDate: t.dueDateTime?.dateTime,
              source: 'ToDo',
              listName: list.displayName,
              webUrl: t.webUrl,
              isCompleted: false
            });
          }
        }
      }

      const plannerRes = await msGraphClient
        .api('/me/planner/tasks')
        .get() as { value: Array<{ id: string; title: string; planId: string; dueDateTime?: string; percentComplete?: number; webUrl?: string }> };

      for (const t of plannerRes.value || []) {
        if (t.percentComplete !== 100) {
          all.push({
            id: `planner-${t.id}`,
            taskId: t.id,
            planId: t.planId,
            title: t.title,
            dueDate: t.dueDateTime,
            source: 'Planner',
            webUrl: t.webUrl,
            isCompleted: false
          });
        }
      }

      all.sort((a, b) => {
        if (!a.dueDate) return 1;
        if (!b.dueDate) return -1;
        return new Date(a.dueDate).getTime() - new Date(b.dueDate).getTime();
      });

      setState({ tasks: all, loading: false });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(`Tasks: ${message}`);
      setState(prev => ({ ...prev, loading: false }));
    }
  }, [msGraphClient, onError]);

  React.useEffect(() => {
    void loadTasks();
  }, [loadTasks]);

  const onComplete = async (task: IUnifiedTask): Promise<void> => {
    try {
      if (task.source === 'ToDo' && task.listId) {
        await msGraphClient
          .api(`/me/todo/lists/${task.listId}/tasks/${task.taskId}`)
          .patch({ status: 'completed' });
      } else if (task.source === 'Planner') {
        const existing = await msGraphClient
          .api(`/me/planner/tasks/${task.taskId}`)
          .select('@odata.etag,percentComplete')
          .get() as { '@odata.etag'?: string };
        const etag = existing['@odata.etag'];
        await msGraphClient
          .api(`/me/planner/tasks/${task.taskId}`)
          .header('If-Match', etag || '*')
          .patch({ percentComplete: 100 });
      }
      setState(prev => ({
        ...prev,
        tasks: prev.tasks.filter(t => t.id !== task.id)
      }));
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(`Complete task: ${message}`);
    }
  };

  if (state.loading) {
    return <Spinner label="Loading tasks..." />;
  }

  if (state.tasks.length === 0) {
    return <p>No pending tasks.</p>;
  }

  return (
    <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
      {state.tasks.map(task => (
        <li key={task.id} style={{ display: 'flex', alignItems: 'center', gap: '8px', marginBottom: '8px' }}>
          <span style={{ flex: 1 }}>
            {task.webUrl ? (
              <a href={task.webUrl} target="_blank" rel="noreferrer">{task.title}</a>
            ) : (
              task.title
            )}
            {task.dueDate && (
              <span style={{ marginLeft: '8px', fontSize: '12px', color: '#605e5c' }}>
                Due {new Date(task.dueDate).toLocaleDateString()}
              </span>
            )}
            <span style={{ marginLeft: '8px', fontSize: '12px' }}>({task.source}{task.listName ? ` · ${task.listName}` : ''})</span>
          </span>
          <PrimaryButton text="Complete" onClick={() => onComplete(task)} />
        </li>
      ))}
    </ul>
  );
};
