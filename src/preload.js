const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('planit', {
  app: {
    version: () => ipcRenderer.invoke('app:version'),
    checkUpdates: () => ipcRenderer.invoke('app:checkUpdates')
  },
  settings: {
    get: () => ipcRenderer.invoke('settings:get'),
    set: (s) => ipcRenderer.invoke('settings:set', s)
  },
  store: {
    get: (name, fallback) => ipcRenderer.invoke('store:get', name, fallback),
    set: (name, data) => ipcRenderer.invoke('store:set', name, data)
  },
  auth: {
    signIn: () => ipcRenderer.invoke('auth:signIn'),
    signOut: () => ipcRenderer.invoke('auth:signOut'),
    account: () => ipcRenderer.invoke('auth:account')
  },
  graph: {
    me: () => ipcRenderer.invoke('graph:me'),
    calendarView: (s, e) => ipcRenderer.invoke('graph:calendarView', s, e),
    createEvent: (ev) => ipcRenderer.invoke('graph:createEvent', ev),
    updateEvent: (id, patch) => ipcRenderer.invoke('graph:updateEvent', id, patch),
    deleteEvent: (id) => ipcRenderer.invoke('graph:deleteEvent', id),
    getSchedule: (emails, s, e, i) => ipcRenderer.invoke('graph:getSchedule', emails, s, e, i),
    listCalendars: () => ipcRenderer.invoke('graph:listCalendars'),
    calendarViewIn: (calId, s, e) => ipcRenderer.invoke('graph:calendarViewIn', calId, s, e),
    shareCalendar: (email, role) => ipcRenderer.invoke('graph:shareCalendar', email, role),
    createCalendar: (name) => ipcRenderer.invoke('graph:createCalendar', name),
    createEventIn: (calId, ev) => ipcRenderer.invoke('graph:createEventIn', calId, ev),
    shareCalendarIn: (calId, email, role) => ipcRenderer.invoke('graph:shareCalendarIn', calId, email, role),
    listMail: (top, skip) => ipcRenderer.invoke('graph:listMail', top, skip),
    readMail: (id) => ipcRenderer.invoke('graph:readMail', id),
    sendMail: (m) => ipcRenderer.invoke('graph:sendMail', m)
  }
});
