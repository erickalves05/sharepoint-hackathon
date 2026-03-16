import * as React from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

export interface IRecentTabProps {
  msGraphClient: MSGraphClientV3;
  onError?: (message: string) => void;
}

interface ITrendingItem {
  id: string;
  resourceVisualization?: {
    title?: string;
    containerDisplayName?: string;
    containerWebUrl?: string;
    previewImageUrl?: string;
  };
  resourceReference?: {
    webUrl?: string;
  };
}

interface IRecentTabState {
  items: ITrendingItem[];
  loading: boolean;
  error?: string;
}

export const RecentTab: React.FunctionComponent<IRecentTabProps> = (props) => {
  const { msGraphClient, onError } = props;
  const [state, setState] = React.useState<IRecentTabState>({ items: [], loading: true });

  React.useEffect(() => {
    let cancelled = false;
    void (async (): Promise<void> => {
      try {
        const result = await msGraphClient
          .api('/me/drive/recent?$top=15')
          .get() as { value?: ITrendingItem[] };
        if (!cancelled) {
          setState({ items: result.value || [], loading: false });
        }
      } catch (err) {
        const message = err instanceof Error ? err.message : String(err);
        onError?.(`Recent: ${message}`);
        if (!cancelled) {
          setState(prev => ({ ...prev, loading: false, error: message }));
        }
      }
    })();
    return (): void => { cancelled = true; };
  }, [msGraphClient, onError]);

  if (state.loading) {
    return <Spinner label="Loading recent files..." />;
  }

  if (state.error) {
    return (
      <MessageBar messageBarType={MessageBarType.warning}>
        {state.error} (Check Sites.Read.All permission and item insights settings.)
      </MessageBar>
    );
  }

  if (state.items.length === 0) {
    return <p>No trending files. Item insights may be disabled for your account.</p>;
  }

  return (
    <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
      {state.items.map(item => {
        const viz = item.resourceVisualization;
        const ref = item.resourceReference;
        const url = ref?.webUrl || viz?.containerWebUrl || '#';
        const title = viz?.title || 'File';
        const location = viz?.containerDisplayName || '';
        return (
          <li key={item.id} style={{ marginBottom: '12px', display: 'flex', alignItems: 'center', gap: '8px' }}>
            {viz?.previewImageUrl && (
              <img src={viz.previewImageUrl} alt="" style={{ width: 32, height: 32, objectFit: 'cover' }} />
            )}
            <span style={{ flex: 1 }}>
              <a href={url} target="_blank" rel="noreferrer">{title}</a>
              {location && <span style={{ marginLeft: '8px', fontSize: '12px', color: '#605e5c' }}>{location}</span>}
            </span>
          </li>
        );
      })}
    </ul>
  );
};
