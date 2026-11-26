const { app, BrowserWindow, shell, Menu } = require('electron');
const APP_URL = 'https://shinesdags-bot.github.io/';
const HOST = 'shinesdags-bot.github.io';

app.commandLine.appendSwitch('autoplay-policy', 'no-user-gesture-required');
app.commandLine.appendSwitch('disable-renderer-backgrounding');
app.setAppUserModelId('io.github.shinesdagsbot.neska');

if (!app.requestSingleInstanceLock()) { app.quit(); }
let win;
app.on('second-instance', () => { if (win) { if (win.isMinimized()) win.restore(); win.focus(); } });

function create() {
  Menu.setApplicationMenu(null);
  win = new BrowserWindow({
    width: 1280, height: 820, minWidth: 380, minHeight: 600,
    title: 'Neska Music', backgroundColor: '#1a0d1f', autoHideMenuBar: true,
    webPreferences: { backgroundThrottling: false, contextIsolation: true, sandbox: true }
  });
  win.loadURL(APP_URL);
  const ext = (url) => { try { return new URL(url).host !== HOST; } catch { return false; } };
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (ext(url) && !/youtube\.com|supabase\.co|google\.com/.test(url)) { shell.openExternal(url); return { action: 'deny' }; }
    return { action: 'allow' };
  });
  win.webContents.on('before-input-event', (e, i) => {
    if (i.type === 'keyDown' && (i.key === 'F5' || (i.control && i.key.toLowerCase() === 'r'))) { win.reload(); e.preventDefault(); }
    if (i.type === 'keyDown' && i.key === 'F11') { win.setFullScreen(!win.isFullScreen()); e.preventDefault(); }
  });
  win.webContents.on('did-fail-load', (_e, code, desc, url, isMain) => {
    if (isMain) win.loadURL('data:text/html,' + encodeURIComponent('<body style="background:#1a0d1f;color:#f3d3c4;font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;flex-direction:column"><h2>Neska Music</h2><p>Can\'t reach the internet. Check your connection.</p><button onclick="location.href=\'' + APP_URL + '\'" style="padding:10px 18px;border-radius:20px;border:0;background:#c98b7a;color:#fff">Try again</button></body>'));
  });
}
app.whenReady().then(create);
app.on('window-all-closed', () => app.quit());
