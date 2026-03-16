import * as React from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import type { ICheckboxStyleProps } from '@fluentui/react/lib/Checkbox';
import type { IUnifiedTask } from './tasks/ITaskItem';

export interface ITasksTabProps {
  msGraphClient: MSGraphClientV3;
  onError?: (message: string) => void;
}

interface ITasksTabState {
  tasks: IUnifiedTask[];
  loading: boolean;
}

function isDueToday(dateStr: string | undefined): boolean {
  if (!dateStr) return false;
  const d = new Date(dateStr);
  const today = new Date();
  return d.getDate() === today.getDate() && d.getMonth() === today.getMonth() && d.getFullYear() === today.getFullYear();
}

export const TasksTab: React.FC<ITasksTabProps> = (props) => {
  const [state, setState] = React.useState<ITasksTabState>({ tasks: [], loading: true });
  const { msGraphClient, onError } = props;

  const loadTasks = React.useCallback(async () => {
    setState(prev => ({ ...prev, loading: true }));
    try {
      const all: IUnifiedTask[] = [];

      const todoLists = await msGraphClient
        .api('/me/todo/lists')
        .get() as { value: Array<{ id: string; displayName: string }> };

      type TodoTaskResponse = {
        id: string;
        title: string;
        status: string;
        dueDateTime?: { dateTime: string };
        webUrl?: string;
        linkedResources?: Array<{ webUrl?: string }>;
      };
      for (const list of todoLists.value || []) {
        const tasksRes = await msGraphClient
          .api(`/me/todo/lists/${list.id}/tasks?$expand=linkedResources`)
          .get() as { value: TodoTaskResponse[] };
        for (const t of tasksRes.value || []) {
          if (t.status !== 'completed') {
            const mailUrl = t.linkedResources?.[0]?.webUrl;
            const displayTitle = t.title.replace(/\s*\(.*$/, '').trim();
            all.push({
              id: t.id,
              taskId: t.id,
              listId: list.id,
              title: displayTitle,
              dueDate: t.dueDateTime?.dateTime,
              source: 'ToDo',
              listName: list.displayName,
              webUrl: mailUrl ?? t.webUrl ?? `https://to-do.office.com/tasks/id/${t.id}/details`,
              isCompleted: false,
              isFlaggedEmail: !!mailUrl
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
            webUrl: `https://planner.cloud.microsoft/webui/plan/${t.planId}/view/board/task/${t.id}`,
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
          .api(`/planner/tasks/${task.taskId}`)
          .get() as { '@odata.etag'?: string };
        const etag = existing['@odata.etag'];
        await msGraphClient
          .api(`/planner/tasks/${task.taskId}`)
          .header('If-Match', etag || '*')
          .patch({ percentComplete: 100 });
      }
      setState(prev => ({
        ...prev,
        tasks: prev.tasks.map(t => t.id === task.id ? { ...t, isCompleted: true } : t)
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
        <li
          key={task.id}
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: '12px',
            marginBottom: '12px',
            padding: '12px',
            border: '1px solid #edebe9',
            borderRadius: '4px',
            backgroundColor: task.isCompleted ? '#faf9f8' : undefined
          }}
        >
          <Checkbox
            ariaLabel={`Mark "${task.title}" as complete`}
            checked={task.isCompleted}
            disabled={task.isCompleted}
            onChange={() => !task.isCompleted && void onComplete(task)}
            styles={{
              root: { margin: 0 },
              checkbox: (styleProps: ICheckboxStyleProps) => ({
                borderRadius: '50%',
                width: 22,
                height: 22,
                ...(styleProps.checked && {
                  background: '#107c10',
                  borderColor: '#107c10'
                })
              }),
              checkmark: (styleProps: ICheckboxStyleProps) =>
                styleProps.checked ? { color: '#fff' } : {}
            }}
          />
          {task.webUrl ? (
            <a
              href={task.webUrl}
              target="_blank"
              rel="noreferrer"
              title={task.isFlaggedEmail ? 'Open mail' : task.source === 'Planner' ? 'Open in Planner' : 'Open in To Do'}
              style={{ display: 'flex', color: '#605e5c', textDecoration: 'none' }}
            >
              <Icon
                iconName={task.isFlaggedEmail ? 'OutlookLogo' : task.source === 'Planner' ? 'PlannerLogo' : 'ToDoLogoOutline'}
                styles={{ root: { fontSize: 18, flexShrink: 0 } }}
              />
            </a>
          ) : (
            <Icon
              iconName={task.isFlaggedEmail ? 'OutlookLogo' : task.source === 'Planner' ? 'PlannerLogo' : 'ToDoLogoOutline'}
              title={task.isFlaggedEmail ? 'Outlook' : task.source === 'Planner' ? 'Microsoft Planner' : 'Microsoft To Do'}
              styles={{ root: { fontSize: 18, color: '#605e5c', flexShrink: 0 } }}
            />
          )}
          <span style={{ flex: 1, minWidth: 0 }}>
            {task.webUrl ? (
              <a
                href={task.webUrl}
                target="_blank"
                rel="noreferrer"
                style={{
                  textDecoration: task.isCompleted ? 'line-through' : undefined,
                  color: task.isCompleted ? '#605e5c' : undefined
                }}
              >
                {task.title}
              </a>
            ) : (
              <span
                style={{
                  textDecoration: task.isCompleted ? 'line-through' : undefined,
                  color: task.isCompleted ? '#605e5c' : undefined
                }}
              >
                {task.title}
              </span>
            )}
            {task.listName && task.source === 'ToDo' && (
              <span style={{ marginLeft: '6px', fontSize: '12px', color: '#605e5c' }}>· {task.listName}</span>
            )}
            {task.dueDate && (
              <span style={{ display: 'inline-flex', alignItems: 'center', marginLeft: '8px', gap: '4px' }}>
                <Icon iconName="Calendar" styles={{ root: { fontSize: 12, color: '#605e5c' } }} />
                {isDueToday(task.dueDate) ? (
                  <span
                    style={{
                      fontSize: '12px',
                      fontWeight: 600,
                      color: '#c43e1c',
                      backgroundColor: '#fde7e6',
                      padding: '2px 8px',
                      borderRadius: '4px'
                    }}
                  >
                    Today
                  </span>
                ) : (
                  <span style={{ fontSize: '12px', color: '#605e5c' }}>
                    {new Date(task.dueDate).toLocaleDateString()}
                  </span>
                )}
              </span>
            )}
          </span>
        </li>
      ))}
    </ul>
  );
};
