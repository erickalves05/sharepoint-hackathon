export type TaskSource = 'ToDo' | 'Planner';

export interface IUnifiedTask {
  id: string;
  title: string;
  dueDate: string | undefined;
  source: TaskSource;
  listName?: string;
  planId?: string;
  webUrl?: string;
  isCompleted: boolean;
  /** For To Do: list id. For Planner: task id. */
  listId?: string;
  /** Needed for PATCH to complete (To Do: listId + taskId; Planner: taskId). */
  taskId: string;
  /** True when task is a flagged email (To Do with linkedResources[0].webUrl). */
  isFlaggedEmail?: boolean;
}
