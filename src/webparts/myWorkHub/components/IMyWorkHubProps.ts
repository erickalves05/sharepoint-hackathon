import type { MSGraphClientV3 } from '@microsoft/sp-http';

/**
 * Options for calling Graph beta API
 */
export interface IGraphBetaRequestOptions {
  method?: 'GET' | 'POST' | 'PATCH' | 'DELETE';
  body?: object;
}

/**
 * Call Graph beta endpoint. path should start with / e.g. /solutions/approval/approvalItems
 */
export type GraphBetaCall = (
  path: string,
  options?: IGraphBetaRequestOptions
) => Promise<{ value?: unknown[]; [key: string]: unknown }>;

export interface IMyWorkHubProps {
  title?: string;
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  /** Microsoft Graph client (v1.0). May be undefined until resolved. */
  msGraphClient?: MSGraphClientV3;
  /** Call Graph beta API. Used for Approvals app. */
  callGraphBeta?: GraphBetaCall;
  /** Current error message to display (set by child via onError). */
  errorMessage?: string | undefined;
  /** Report an error to show in the web part (e.g. permission or API failure). */
  onError?: (message: string) => void;
  /** Clear the current error message. */
  dismissError?: () => void;
}
