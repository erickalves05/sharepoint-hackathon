import * as React from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IRecentTabProps {
  msGraphClient: MSGraphClientV3;
  onError?: (message: string) => void;
}

interface IDriveItem {
  id: string;
  name: string;
  webUrl?: string;
  lastModifiedDateTime?: string;
  createdBy?: { user?: { displayName?: string } };
  lastModifiedBy?: { user?: { displayName?: string } };
  file?: { mimeType?: string };
  remoteItem?: {
    webUrl?: string;
    name?: string;
    lastModifiedDateTime?: string;
    createdBy?: { user?: { displayName?: string } };
    lastModifiedBy?: { user?: { displayName?: string } };
    file?: { mimeType?: string };
  };
}

interface IRecentTabState {
  items: IDriveItem[];
  loading: boolean;
  error?: string;
}

function getFileIcon(item: IDriveItem): string {
  const mime = item.file?.mimeType ?? item.remoteItem?.file?.mimeType ?? '';
  const ext = (item.name ?? '').split('.').pop()?.toLowerCase() ?? '';
  if (mime.includes('spreadsheet') || mime.includes('ms-excel') || ext === 'xlsx' || ext === 'xls') return 'ExcelDocument';
  if (mime.includes('wordprocessing') || mime.includes('msword') || ext === 'docx' || ext === 'doc') return 'WordDocument';
  if (mime.includes('presentation') || mime.includes('ms-powerpoint') || ext === 'pptx' || ext === 'ppt') return 'PowerPointDocument';
  if (mime === 'application/pdf' || ext === 'pdf') return 'PDF';
  if (mime.includes('onenote') || ext === 'one') return 'OneNoteLogo';
  return 'Page';
}

export const RecentTab: React.FC<IRecentTabProps> = (props) => {
  const { msGraphClient, onError } = props;
  const [state, setState] = React.useState<IRecentTabState>({ items: [], loading: true });

  React.useEffect(() => {
    let cancelled = false;
    void (async (): Promise<void> => {
      try {
        const result = await msGraphClient
          .api('/me/drive/recent?$top=15')
          .get() as { value?: IDriveItem[] };
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
        {state.error} (Check Files.Read.All permission.)
      </MessageBar>
    );
  }

  if (state.items.length === 0) {
    return <p>No recent files.</p>;
  }

  return (
    <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
      {state.items.map(item => {
        const url = item.webUrl ?? item.remoteItem?.webUrl ?? '#';
        const name = item.name ?? item.remoteItem?.name ?? 'File';
        const lastModified = item.lastModifiedDateTime ?? item.remoteItem?.lastModifiedDateTime;
        const creator = "Erick Alves"//item.createdBy?.user?.displayName ?? item.remoteItem?.createdBy?.user?.displayName
          //?? item.lastModifiedBy?.user?.displayName ?? item.remoteItem?.lastModifiedBy?.user?.displayName;
        const content = (
          <>
            <Icon
              iconName={getFileIcon(item)}
              styles={{ root: { fontSize: 28, color: '#605e5c', flexShrink: 0 } }}
            />
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontWeight: 600 }}>{name}</div>
              <div style={{ fontSize: '12px', color: '#605e5c', marginTop: '4px' }}>
                {creator && <span>{creator}</span>}
                {creator && lastModified && <span> · </span>}
                {lastModified && (
                  <span>{new Date(lastModified).toLocaleString()}</span>
                )}
              </div>
            </div>
          </>
        );
        return (
          <li key={item.id} style={{ marginBottom: '12px' }}>
            {url !== '#' ? (
              <a
                href={url}
                target="_blank"
                rel="noreferrer"
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '12px',
                  padding: '12px',
                  border: '1px solid #edebe9',
                  borderRadius: '4px',
                  textDecoration: 'none',
                  color: 'inherit',
                  cursor: 'pointer'
                }}
              >
                {content}
              </a>
            ) : (
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '12px', border: '1px solid #edebe9', borderRadius: '4px' }}>
                {content}
              </div>
            )}
          </li>
        );
      })}
    </ul>
  );
};
