// Planit — Electron main process
const { app, BrowserWindow, ipcMain, shell, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

const auth = require('./auth');
const graph = require('./graph');

let win = null;

// ---------- simple JSON stores in userData ----------
function storePath(name) {
  return path.join(app.getPath('userData'), name + '.json');
}
function readStore(name, fallback) {
  try {
    return JSON.parse(fs.readFileSync(storePath(name), 'utf8'));
  } catch {
    return fallback;
  }
}
function writeStore(name, data) {
  fs.mkdirSync(app.getPath('userData'), { recursive: true });
  fs.writeFileSync(storePath(name), JSON.stringify(data, null, 2), 'utf8');
}

function createWindow() {
  win = new BrowserWindow({
    width: 1360,
    height: 860,
    minWidth: 980,
    minHeight: 640,
    backgroundColor: '#f6f9fb',
    title: 'Planit',
    icon: path.join(__dirname, '..', 'assets', 'icon.png'),
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  });
  win.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));

  // open external links in the system browser
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('http')) shell.openExternal(url);
    return { action: 'deny' };
  });
}

// ---------- auto-update (GitHub Releases) ----------
function setupAutoUpdate() {
  if (!app.isPackaged) return; // dev runs never auto-update
  try {
    const { autoUpdater } = require('electron-updater');
    autoUpdater.autoDownload = true;
    autoUpdater.on('update-downloaded', (info) => {
      dialog.showMessageBox(win, {
        type: 'info',
        buttons: ['Restart now', 'Later'],
        defaultId: 0,
        title: 'Planit update',
        message: `Planit ${info.version} is ready to install.`,
        detail: 'The update was downloaded in the background. Restart Planit to apply it.'
      }).then(r => { if (r.response === 0) autoUpdater.quitAndInstall(); });
    });
    autoUpdater.on('error', () => { /* offline or no release yet — stay quiet */ });
    autoUpdater.checkForUpdatesAndNotify().catch(() => {});
    // re-check every 4 hours while the app stays open
    setInterval(() => autoUpdater.checkForUpdates().catch(() => {}), 4 * 60 * 60 * 1000);
  } catch { /* updater not available */ }
}

app.whenReady().then(() => {
  createWindow();
  setupAutoUpdate();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

ipcMain.handle('app:version', () => app.getVersion());
ipcMain.handle('app:checkUpdates', async () => {
  if (!app.isPackaged) return { ok: false, reason: 'dev' };
  try {
    const { autoUpdater } = require('electron-updater');
    const res = await autoUpdater.checkForUpdates();
    const latest = res && res.updateInfo && res.updateInfo.version;
    return { ok: true, current: app.getVersion(), latest };
  } catch (e) {
    return { ok: false, reason: String(e && e.message || e) };
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ---------- settings ----------
ipcMain.handle('settings:get', () => {
  return readStore('settings', { clientId: '', tenant: 'common', demoMode: true });
});
ipcMain.handle('settings:set', (e, settings) => {
  writeStore('settings', settings);
  auth.configure(settings, app.getPath('userData'));
  return true;
});

// ---------- local data (polls, local events fallback) ----------
ipcMain.handle('store:get', (e, name, fallback) => readStore(name, fallback));
ipcMain.handle('store:set', (e, name, data) => {
  writeStore(name, data);
  return true;
});

// ---------- auth ----------
ipcMain.handle('auth:signIn', async () => {
  const settings = readStore('settings', {});
  auth.configure(settings, app.getPath('userData'));
  return auth.signIn();
});
ipcMain.handle('auth:signOut', async () => auth.signOut());
ipcMain.handle('auth:account', async () => {
  const settings = readStore('settings', {});
  auth.configure(settings, app.getPath('userData'));
  return auth.getAccount();
});

// ---------- Microsoft Graph ----------
async function withToken(fn) {
  const token = await auth.getToken();
  return fn(token);
}

ipcMain.handle('graph:me', () => withToken(t => graph.me(t)));
ipcMain.handle('graph:calendarView', (e, startISO, endISO) =>
  withToken(t => graph.calendarView(t, startISO, endISO)));
ipcMain.handle('graph:createEvent', (e, event) =>
  withToken(t => graph.createEvent(t, event)));
ipcMain.handle('graph:updateEvent', (e, id, patch) =>
  withToken(t => graph.updateEvent(t, id, patch)));
ipcMain.handle('graph:deleteEvent', (e, id) =>
  withToken(t => graph.deleteEvent(t, id)));
ipcMain.handle('graph:getSchedule', (e, emails, startISO, endISO, intervalMin) =>
  withToken(t => graph.getSchedule(t, emails, startISO, endISO, intervalMin)));
ipcMain.handle('graph:listCalendars', () => withToken(t => graph.listCalendars(t)));
ipcMain.handle('graph:calendarViewIn', (e, calId, startISO, endISO) =>
  withToken(t => graph.calendarViewIn(t, calId, startISO, endISO)));
ipcMain.handle('graph:shareCalendar', (e, email, role) =>
  withToken(t => graph.shareCalendar(t, email, role)));
ipcMain.handle('graph:createCalendar', (e, name) =>
  withToken(t => graph.createCalendar(t, name)));
ipcMain.handle('graph:createEventIn', (e, calId, event) =>
  withToken(t => graph.createEventIn(t, calId, event)));
ipcMain.handle('graph:shareCalendarIn', (e, calId, email, role) =>
  withToken(t => graph.shareCalendarIn(t, calId, email, role)));
ipcMain.handle('graph:listMail', (e, top, skip) =>
  withToken(t => graph.listMail(t, top, skip)));
ipcMain.handle('graph:readMail', (e, id) => withToken(t => graph.readMail(t, id)));
ipcMain.handle('graph:sendMail', (e, message) =>
  withToken(t => graph.sendMail(t, message)));
