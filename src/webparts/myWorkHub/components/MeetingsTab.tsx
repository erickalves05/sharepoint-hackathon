import * as React from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import { Spinner } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IMeetingsTabProps {
  msGraphClient: MSGraphClientV3;
  onError?: (message: string) => void;
}

interface IMeetingItem {
  id: string;
  subject: string;
  startDateTime: string;
  endDateTime: string;
  webLink?: string;
  organizer?: string;
}

interface IMeetingsTabState {
  meetings: IMeetingItem[];
  loading: boolean;
  error?: string;
}

function isDueToday(dateStr: string | undefined): boolean {
  if (!dateStr) return false;
  const d = new Date(dateStr);
  const today = new Date();
  return d.getDate() === today.getDate() && d.getMonth() === today.getMonth() && d.getFullYear() === today.getFullYear();
}

function isDueTomorrow(dateStr: string | undefined): boolean {
  if (!dateStr) return false;
  const d = new Date(dateStr);
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  return d.getDate() === tomorrow.getDate() && d.getMonth() === tomorrow.getMonth() && d.getFullYear() === tomorrow.getFullYear();
}

function formatMeetingDate(dateStr: string): string {
  if (isDueToday(dateStr)) return 'Today';
  if (isDueTomorrow(dateStr)) return 'Tomorrow';
  return new Date(dateStr).toLocaleDateString();
}

function formatTime(dateStr: string): string {
  return new Date(dateStr).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
}

export const MeetingsTab: React.FC<IMeetingsTabProps> = (props) => {
  const { msGraphClient, onError } = props;
  const [state, setState] = React.useState<IMeetingsTabState>({ meetings: [], loading: true });

  const loadMeetings = React.useCallback(async () => {
    setState(prev => ({ ...prev, loading: true, error: undefined }));
    try {
      const now = new Date().toISOString();
      const result = await msGraphClient
        .api(`/me/events?$filter=start/dateTime ge '${now}'&$orderby=start/dateTime&$top=5&$select=subject,start,end,organizer,webLink`)
        .get() as { value?: Array<{
          id: string;
          subject?: string;
          start?: { dateTime: string };
          end?: { dateTime: string };
          webLink?: string;
          organizer?: { emailAddress?: { name?: string } };
        }> };

      const meetings: IMeetingItem[] = (result.value || []).map((e) => ({
        id: e.id,
        subject: e.subject ?? 'No subject',
        startDateTime: e.start?.dateTime ?? '',
        endDateTime: e.end?.dateTime ?? '',
        webLink: e.webLink,
        organizer: "Erick Alves" //e.organizer?.emailAddress?.name
      }));

      setState({ meetings, loading: false });
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      onError?.(`Meetings: ${message}`);
      setState(prev => ({ ...prev, loading: false, error: message }));
    }
  }, [msGraphClient, onError]);

  React.useEffect(() => {
    void loadMeetings();
  }, [loadMeetings]);

  if (state.loading) {
    return <Spinner label="Loading meetings..." />;
  }

  if (state.error) {
    return (
      <p style={{ color: '#c43e1c' }}>
        {state.error} (Check Calendars.Read permission.)
      </p>
    );
  }

  if (state.meetings.length === 0) {
    return <p>No upcoming meetings.</p>;
  }

  return (
    <ul style={{ listStyle: 'none', padding: 0, margin: 0 }}>
      {state.meetings.map((meeting) => {
        const content = (
          <>
            <Icon
              iconName="Calendar"
              styles={{ root: { fontSize: 24, color: '#605e5c', flexShrink: 0 } }}
            />
            <span style={{ flex: 1, minWidth: 0 }}>
              <span style={{ fontWeight: 600 }}>{meeting.subject}</span>
              {meeting.organizer && (
                <span style={{ marginLeft: '6px', fontSize: '12px', color: '#605e5c' }}>· {meeting.organizer}</span>
              )}
              <span style={{ display: 'inline-flex', alignItems: 'center', marginLeft: '8px', gap: '4px', flexWrap: 'wrap' }}>
                {meeting.startDateTime && (
                  isDueToday(meeting.startDateTime) ? (
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
                  ) : isDueTomorrow(meeting.startDateTime) ? (
                    <span
                      style={{
                        fontSize: '12px',
                        fontWeight: 600,
                        color: '#106ebe',
                        backgroundColor: '#e6f2fa',
                        padding: '2px 8px',
                        borderRadius: '4px'
                      }}
                    >
                      Tomorrow
                    </span>
                  ) : (
                    <span style={{ fontSize: '12px', color: '#605e5c' }}>
                      {formatMeetingDate(meeting.startDateTime)}
                    </span>
                  )
                )}
                <span style={{ fontSize: '12px', color: '#605e5c' }}>
                  {formatTime(meeting.startDateTime)} – {formatTime(meeting.endDateTime)}
                </span>
              </span>
            </span>
          </>
        );
        return (
          <li key={meeting.id} style={{ marginBottom: '12px' }}>
            {meeting.webLink ? (
              <a
                href={meeting.webLink}
                target="_blank"
                rel="noreferrer"
                style={{
                  display: 'flex',
                  alignItems: 'flex-start',
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
              <div
                style={{
                  display: 'flex',
                  alignItems: 'flex-start',
                  gap: '12px',
                  padding: '12px',
                  border: '1px solid #edebe9',
                  borderRadius: '4px'
                }}
              >
                {content}
              </div>
            )}
          </li>
        );
      })}
    </ul>
  );
};
