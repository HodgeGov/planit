// Thin Microsoft Graph REST wrappers (Node 18+ global fetch).
const BASE = 'https://graph.microsoft.com/v1.0';

async function call(token, method, url, body, extraHeaders) {
  const res = await fetch(BASE + url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...(extraHeaders || {})
    },
    body: body ? JSON.stringify(body) : undefined
  });
  if (res.status === 204) return null;
  const text = await res.text();
  const data = text ? JSON.parse(text) : null;
  if (!res.ok) {
    const msg = (data && data.error && data.error.message) || `Graph ${res.status}`;
    const err = new Error(msg);
    err.status = res.status;
    throw err;
  }
  return data;
}

const me = (t) => call(t, 'GET', '/me');

// Calendar ---------------------------------------------------------------
const calendarView = (t, startISO, endISO) =>
  call(t, 'GET',
    `/me/calendarView?startDateTime=${encodeURIComponent(startISO)}` +
    `&endDateTime=${encodeURIComponent(endISO)}` +
    `&$top=250&$orderby=start/dateTime` +
    `&$select=id,subject,start,end,isAllDay,location,organizer,attendees,bodyPreview,isOnlineMeeting,onlineMeeting`,
    null,
    { Prefer: 'outlook.timezone="UTC"' });

const createEvent = (t, event) => call(t, 'POST', '/me/events', event);

// All calendars in the mailbox — includes calendars other people shared with you.
const listCalendars = (t) =>
  call(t, 'GET', '/me/calendars?$top=50&$select=id,name,owner,canEdit,canShare,isDefaultCalendar');

// Create a new (group) calendar in your mailbox.
const createCalendar = (t, name) => call(t, 'POST', '/me/calendars', { name });

// Create an event inside a specific calendar (e.g. a group calendar).
const createEventIn = (t, calId, event) =>
  call(t, 'POST', `/me/calendars/${calId}/events`, event);

// Grant someone access to a specific calendar ('read' or 'write' — write lets them plan events).
const shareCalendarIn = (t, calId, email, role) =>
  call(t, 'POST', `/me/calendars/${calId}/calendarPermissions`, {
    emailAddress: { address: email },
    role: role || 'read',
    allowedRoles: role === 'write' ? ['read', 'write'] : ['read']
  });

// Events from a specific (e.g. shared) calendar.
const calendarViewIn = (t, calId, startISO, endISO) =>
  call(t, 'GET',
    `/me/calendars/${calId}/calendarView?startDateTime=${encodeURIComponent(startISO)}` +
    `&endDateTime=${encodeURIComponent(endISO)}` +
    `&$top=100&$orderby=start/dateTime` +
    `&$select=id,subject,start,end,isAllDay,location,organizer,isOnlineMeeting,onlineMeeting`,
    null,
    { Prefer: 'outlook.timezone="UTC"' });

// Grant someone read access to your default calendar (sends an Outlook sharing invite).
const shareCalendar = (t, email, role) =>
  call(t, 'POST', '/me/calendar/calendarPermissions', {
    emailAddress: { address: email },
    role: role || 'read',
    allowedRoles: ['read']
  });
const updateEvent = (t, id, patch) => call(t, 'PATCH', `/me/events/${id}`, patch);
const deleteEvent = (t, id) => call(t, 'DELETE', `/me/events/${id}`);

// Free/busy for a set of people (work/school accounts) -------------------
const getSchedule = (t, emails, startISO, endISO, intervalMin) =>
  call(t, 'POST', '/me/calendar/getSchedule', {
    schedules: emails,
    startTime: { dateTime: startISO, timeZone: 'UTC' },
    endTime: { dateTime: endISO, timeZone: 'UTC' },
    availabilityViewInterval: intervalMin || 30
  });

// Mail -------------------------------------------------------------------
const listMail = (t, top, skip) =>
  call(t, 'GET',
    `/me/mailFolders/inbox/messages?$top=${top || 25}&$skip=${skip || 0}` +
    `&$orderby=receivedDateTime desc` +
    `&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments`);

const readMail = (t, id) =>
  call(t, 'GET', `/me/messages/${id}?$select=id,subject,from,toRecipients,receivedDateTime,body`);

const sendMail = (t, message) =>
  call(t, 'POST', '/me/sendMail', { message, saveToSentItems: true });

module.exports = {
  me, calendarView, createEvent, updateEvent, deleteEvent,
  listCalendars, calendarViewIn, shareCalendar,
  createCalendar, createEventIn, shareCalendarIn,
  getSchedule, listMail, readMail, sendMail
};
