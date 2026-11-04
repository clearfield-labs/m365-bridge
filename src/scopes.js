// Permission packs. Every scope here is DELEGATED: the token acts as the signed-in person
// and can only reach what that person can already reach in Microsoft 365.
// The admin picks packs once; the setup code carries them to the team.
export const PACKS = {
  files: {
    label: 'Files (OneDrive + SharePoint)',
    read: ['Files.Read.All', 'Sites.Read.All'],
    write: ['Files.ReadWrite.All', 'Sites.ReadWrite.All'],
    always: true,
  },
  mail: {
    label: 'Outlook mail (own + shared mailboxes they have rights to)',
    read: ['Mail.Read', 'Mail.Read.Shared', 'MailboxSettings.Read'],
    write: ['Mail.ReadWrite', 'Mail.ReadWrite.Shared', 'Mail.Send', 'Mail.Send.Shared', 'MailboxSettings.ReadWrite'],
  },
  calendar: {
    label: 'Calendar (own + shared calendars)',
    read: ['Calendars.Read', 'Calendars.Read.Shared'],
    write: ['Calendars.ReadWrite', 'Calendars.ReadWrite.Shared'],
  },
  contacts: {
    label: 'Contacts',
    read: ['Contacts.Read', 'Contacts.Read.Shared'],
    write: ['Contacts.ReadWrite', 'Contacts.ReadWrite.Shared'],
  },
  teams: {
    label: 'Teams chats and channels',
    read: ['Chat.Read', 'ChannelMessage.Read.All', 'Team.ReadBasic.All', 'Channel.ReadBasic.All'],
    write: ['Chat.ReadWrite', 'ChannelMessage.Read.All', 'ChannelMessage.Send', 'Team.ReadBasic.All', 'Channel.ReadBasic.All'],
  },
  meetings: {
    label: 'Teams meetings, recordings and transcripts',
    read: ['OnlineMeetings.Read', 'OnlineMeetingTranscript.Read.All', 'OnlineMeetingRecording.Read.All'],
    write: ['OnlineMeetings.ReadWrite', 'OnlineMeetingTranscript.Read.All', 'OnlineMeetingRecording.Read.All'],
  },
  tasks: {
    label: 'To Do and Planner tasks',
    read: ['Tasks.Read', 'Tasks.Read.Shared'],
    write: ['Tasks.ReadWrite', 'Tasks.ReadWrite.Shared'],
  },
  notes: {
    label: 'OneNote',
    read: ['Notes.Read.All'],
    write: ['Notes.ReadWrite.All'],
  },
  people: {
    label: 'People, directory and presence',
    read: ['People.Read', 'User.ReadBasic.All', 'Presence.Read.All'],
    write: ['People.Read', 'User.ReadBasic.All', 'Presence.Read.All'],
  },
};

export const DEFAULT_PACKS = ['files', 'mail', 'calendar', 'contacts', 'teams', 'meetings', 'tasks', 'notes', 'people'];

const SCOPE_RE = /^[A-Za-z]+(\.[A-Za-z]+)+$/;

export function parseExtraScopes(text) {
  return String(text || '')
    .split(/[\s,]+/)
    .map((s) => s.trim().replace(/^https:\/\/graph\.microsoft\.com\//i, ''))
    .filter((s) => SCOPE_RE.test(s));
}

export function scopesFor(cfg) {
  const set = new Set(['User.Read']);
  const packs = new Set(['files', ...(cfg.packs || [])]);
  for (const key of packs) {
    const pack = PACKS[key];
    if (pack) for (const s of cfg.allowWrite ? pack.write : pack.read) set.add(s);
  }
  for (const s of parseExtraScopes(cfg.extraScopes)) set.add(s);
  return [...set];
}

export const hasPack = (cfg, key) => key === 'files' || (cfg.packs || []).includes(key);

/**
 * One-click tenant-wide consent for exactly the selected delegated scopes (v2 dynamic consent).
 * The admin does not have to add each permission in the Entra portal first.
 */
export function adminConsentUrl(cfg, redirectUri) {
  const scope = scopesFor(cfg).map((s) => `https://graph.microsoft.com/${s}`).join(' ');
  const q = new URLSearchParams({ client_id: cfg.clientId, scope, redirect_uri: redirectUri, state: 'adminconsent' });
  return `https://login.microsoftonline.com/${cfg.tenantId}/v2.0/adminconsent?${q}`;
}
