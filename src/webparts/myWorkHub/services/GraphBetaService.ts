import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { IGraphBetaRequestOptions } from '../components/IMyWorkHubProps';

const GRAPH_BETA_BASE = 'https://graph.microsoft.com/beta';

function toBetaPath(path: string): string {
  if (path.indexOf('http') === 0) return path;
  const p = path.indexOf('/') === 0 ? path : `/${path}`;
  return `${GRAPH_BETA_BASE}${p}`;
}

/**
 * Create a function that calls Microsoft Graph beta using the same MSGraphClient (same token).
 * Uses client.api(betaUrl) so no separate token is needed for the beta endpoint.
 */
export function createGraphBetaCallFromClient(
  client: MSGraphClientV3
): (path: string, options?: IGraphBetaRequestOptions) => Promise<{ value?: unknown[]; [key: string]: unknown }> {
  return async (
    path: string,
    options?: IGraphBetaRequestOptions
  ): Promise<{ value?: unknown[]; [key: string]: unknown }> => {
    const fullPath = toBetaPath(path);
    const method = options?.method || 'GET';
    if (method === 'POST') {
      return (await client.api(fullPath).post(options?.body ?? {})) as { value?: unknown[]; [key: string]: unknown };
    }
    if (method === 'PATCH') {
      return (await client.api(fullPath).patch(options?.body ?? {})) as { value?: unknown[]; [key: string]: unknown };
    }
    return (await client.api(fullPath).get()) as { value?: unknown[]; [key: string]: unknown };
  };
}
