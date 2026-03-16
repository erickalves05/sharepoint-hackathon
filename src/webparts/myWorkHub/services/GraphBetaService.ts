import type { IGraphBetaRequestOptions } from '../components/IMyWorkHubProps';

const GRAPH_BETA_BASE = 'https://graph.microsoft.com/beta';

/**
 * Create a function that calls Microsoft Graph beta using the provided token getter.
 * Used for Approvals app API (beta only).
 */
export function createGraphBetaCall(
  getToken: () => Promise<string>
): (path: string, options?: IGraphBetaRequestOptions) => Promise<{ value?: unknown[]; [key: string]: unknown }> {
  return async (
    path: string,
    options?: IGraphBetaRequestOptions
  ): Promise<{ value?: unknown[]; [key: string]: unknown }> => {
    const url = path.indexOf('http') === 0 ? path : `${GRAPH_BETA_BASE}${path.indexOf('/') === 0 ? path : `/${path}`}`;
    const token = await getToken();
    const init: RequestInit = {
      method: options?.method || 'GET',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        Accept: 'application/json'
      }
    };
    if (options?.body && (options.method === 'POST' || options.method === 'PATCH')) {
      init.body = JSON.stringify(options.body);
    }
    const response = await fetch(url, init);
    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Graph beta error ${response.status}: ${text || response.statusText}`);
    }
    const json = await response.json();
    return json as { value?: unknown[]; [key: string]: unknown };
  };
}
