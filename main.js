'use strict';

const { app, BrowserWindow, ipcMain, dialog, shell, Menu, session, screen, nativeImage } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');

const APP_NAME = 'Redator';
const TITLE_BAR_COLOR = '#185abd';
const SUPPORTED = ['.docx', '.docm', '.dotx', '.md', '.markdown', '.txt', '.text', '.log', '.csv', '.html', '.htm', '.odt', '.rtf'];

app.setName(APP_NAME);
if (process.platform === 'win32') app.setAppUserModelId('br.redator.app');
if (process.env.REDATOR_TEST) {
  app.setPath('userData', path.join(process.env.REDATOR_TEST_OUT || os.tmpdir(), 'redator-test-profile'));
  // máquinas de CI nem sempre têm placa de vídeo
  app.disableHardwareAcceleration();
}

const userData = app.getPath('userData');
const recentFile = path.join(userData, 'recentes.json');
const windowStateFile = path.join(userData, 'janela.json');
const prefsFile = path.join(userData, 'preferencias.json');

// ---------------------------------------------------------------- utilidades

function readJson(file, fallback) {
  try { return JSON.parse(fs.readFileSync(file, 'utf8')); } catch { return fallback; }
}

function writeJson(file, value) {
  try {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, JSON.stringify(value, null, 2), 'utf8');
  } catch (err) {
    console.error('Falha ao gravar', file, err);
  }
}

function isSupportedFile(p) {
  try {
    if (!p || p.startsWith('--')) return false;
    const st = fs.statSync(p);
    return st.isFile() && SUPPORTED.includes(path.extname(p).toLowerCase());
  } catch { return false; }
}

function filesFromArgv(argv, cwd) {
  const out = [];
  for (const raw of argv.slice(1)) {
    if (!raw || raw.startsWith('-')) continue;
    const p = path.resolve(cwd || process.cwd(), raw);
    if (p === path.resolve(__dirname)) continue;
    if (isSupportedFile(p)) out.push(p);
  }
  return out;
}

// ---------------------------------------------------------------- recentes

function getRecent() {
  const list = readJson(recentFile, []);
  return list.filter((r) => r && r.path);
}

function addRecent(p) {
  if (!p) return getRecent();
  const list = getRecent().filter((r) => r.path.toLowerCase() !== p.toLowerCase());
  list.unshift({ path: p, name: path.basename(p), folder: path.dirname(p), time: Date.now() });
  const trimmed = list.slice(0, 40);
  writeJson(recentFile, trimmed);
  try { app.addRecentDocument(p); } catch { /* ignora */ }
  return trimmed;
}

function removeRecent(p) {
  const list = getRecent().filter((r) => r.path.toLowerCase() !== String(p).toLowerCase());
  writeJson(recentFile, list);
  return list;
}

// ---------------------------------------------------------------- janelas

const windows = new Set();

function nextBounds() {
  const saved = readJson(windowStateFile, null);
  const area = screen.getPrimaryDisplay().workArea;
  let b = saved && saved.width ? saved : { width: Math.min(1400, area.width), height: Math.min(900, area.height), maximized: true };
  if (windows.size > 0) {
    const last = [...windows].pop();
    if (last && !last.isDestroyed()) {
      const lb = last.getBounds();
      b = { ...b, x: lb.x + 28, y: lb.y + 28, width: lb.width, height: lb.height, maximized: last.isMaximized() };
    }
  }
  return b;
}

function saveWindowState(win) {
  if (!win || win.isDestroyed()) return;
  const maximized = win.isMaximized();
  const b = maximized ? readJson(windowStateFile, {}) : win.getBounds();
  writeJson(windowStateFile, { ...b, maximized });
}

function createWindow(filePath) {
  const b = nextBounds();
  const iconPath = path.join(__dirname, 'app', 'icons', 'redator.ico');
  const win = new BrowserWindow({
    x: b.x, y: b.y, width: b.width || 1280, height: b.height || 860,
    minWidth: 560, minHeight: 420,
    show: false,
    backgroundColor: '#e8e8e8',
    title: APP_NAME,
    icon: fs.existsSync(iconPath) ? iconPath : undefined,
    titleBarStyle: 'hidden',
    titleBarOverlay: { color: TITLE_BAR_COLOR, symbolColor: '#ffffff', height: 40 },
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: true,
    },
  });

  win.docState = { filePath: filePath || null, dirty: false, allowClose: false };
  windows.add(win);

  win.loadFile(path.join(__dirname, 'app', 'index.html'));

  win.once('ready-to-show', () => {
    if (b.maximized) win.maximize();
    win.show();
  });

  win.on('close', (e) => {
    saveWindowState(win);
    if (win.docState.dirty && !win.docState.allowClose) {
      e.preventDefault();
      win.webContents.send('app:request-close');
    }
  });

  win.on('closed', () => { windows.delete(win); });

  win.on('maximize', () => win.webContents.send('app:window-state', { maximized: true }));
  win.on('unmaximize', () => win.webContents.send('app:window-state', { maximized: false }));
  win.on('focus', () => win.webContents.send('app:focus'));

  const wc = win.webContents;
  wc.on('will-navigate', (e) => e.preventDefault());
  wc.setWindowOpenHandler(({ url }) => {
    if (/^(https?|mailto):/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  wc.on('context-menu', (_e, params) => {
    if (process.env.REDATOR_TEST) {
      const ses = session.defaultSession;
      console.log('TEST context-menu', JSON.stringify({ w: params.misspelledWord, s: params.dictionarySuggestions, spell: params.spellcheckEnabled, editable: params.isEditable, enabled: ses.spellCheckerEnabled, langs: ses.getSpellCheckerLanguages(), avail: (ses.availableSpellCheckerLanguages || []).slice(0, 80) }));
    }
    wc.send('app:context-menu', {
      x: params.x, y: params.y,
      misspelledWord: params.misspelledWord || '',
      suggestions: params.dictionarySuggestions || [],
      isEditable: params.isEditable,
      selectionText: params.selectionText || '',
      linkURL: params.linkURL || '',
      editFlags: params.editFlags || {},
    });
  });
  if (process.env.REDATOR_TEST) attachTestHarness(win);
  wc.on('before-input-event', (event, input) => {
    if (input.type === 'keyDown' && input.control && input.shift && input.alt && input.key.toLowerCase() === 'i') {
      wc.toggleDevTools();
      event.preventDefault();
    }
  });
  return win;
}

function windowFor(event) {
  return BrowserWindow.fromWebContents(event.sender);
}

function openFileInNewWindow(p) {
  // se já estiver aberto, só foca
  for (const w of windows) {
    if (!w.isDestroyed() && w.docState.filePath && w.docState.filePath.toLowerCase() === p.toLowerCase()) {
      if (w.isMinimized()) w.restore();
      w.focus();
      return w;
    }
  }
  return createWindow(p);
}

// ---------------------------------------------------------------- instância única

const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on('second-instance', (_event, argv, cwd) => {
    const files = filesFromArgv(argv, cwd);
    if (files.length) files.forEach(openFileInNewWindow);
    else createWindow(null);
  });

  app.whenReady().then(() => {
    Menu.setApplicationMenu(null);
    configureSpellchecker();
    if (process.env.REDATOR_TEST) {
      // nos testes locais nada sai para a rede
      session.defaultSession.webRequest.onBeforeRequest({ urls: ['http://*/*', 'https://*/*', 'ws://*/*', 'wss://*/*'] }, (details, cb) => {
        console.log('TEST net-blocked', details.url);
        cb({ cancel: true });
      });
    }
    setTimeout(startSpellHelper, 2500);
    const files = filesFromArgv(process.argv, process.cwd());
    if (files.length) files.forEach((f) => createWindow(f));
    else createWindow(null);
  });
}

app.on('window-all-closed', () => { app.quit(); });

// ---------------------------------------------------------------- ortografia

function configureSpellchecker() {
  const ses = session.defaultSession;
  const prefs = readJson(prefsFile, {});
  const wanted = prefs.spellLanguages || ['pt-BR'];
  try {
    const available = ses.availableSpellCheckerLanguages || [];
    const langs = wanted.filter((l) => !available.length || available.includes(l));
    ses.setSpellCheckerLanguages(langs.length ? langs : wanted);
  } catch (err) {
    console.warn('Idiomas de ortografia indisponíveis', err.message);
  }
  ses.setPermissionRequestHandler((_wc, permission, cb) => {
    cb(permission === 'local-fonts' || permission === 'clipboard-read' || permission === 'clipboard-sanitized-write');
  });
  ses.setPermissionCheckHandler((_wc, permission) => permission === 'local-fonts' || permission === 'clipboard-read' || permission === 'clipboard-sanitized-write');
}

// ---------------------------------------------------------------- IPC

ipcMain.handle('app:init', (event) => {
  const win = windowFor(event);
  return {
    filePath: win ? win.docState.filePath : null,
    recent: getRecent(),
    version: app.getVersion(),
    userName: os.userInfo().username,
    prefs: readJson(prefsFile, {}),
    platform: process.platform,
  };
});

ipcMain.handle('file:read', async (_event, p) => {
  const data = await fs.promises.readFile(p);
  let readonly = false;
  try { await fs.promises.access(p, fs.constants.W_OK); } catch { readonly = true; }
  const st = await fs.promises.stat(p);
  return { data: new Uint8Array(data.buffer, data.byteOffset, data.byteLength), readonly, mtime: st.mtimeMs, size: st.size, name: path.basename(p), dir: path.dirname(p) };
});

ipcMain.handle('file:exists', (_event, p) => {
  try { return fs.statSync(p).isFile(); } catch { return false; }
});

ipcMain.handle('file:write', async (_event, p, data, extras) => {
  const buf = typeof data === 'string' ? Buffer.from(data, 'utf8') : Buffer.from(data);
  const tmp = p + '.' + process.pid + '.redator-tmp';
  try {
    await fs.promises.writeFile(tmp, buf);
    await fs.promises.rename(tmp, p);
  } catch (err) {
    try { await fs.promises.unlink(tmp); } catch { /* ignora */ }
    if (err.code === 'EPERM' || err.code === 'EBUSY' || err.code === 'EACCES') {
      // tenta gravar direto, alguns locais bloqueiam o rename
      try { await fs.promises.writeFile(p, buf); } catch (err2) { return { ok: false, error: err2.code || err2.message }; }
    } else {
      return { ok: false, error: err.code || err.message };
    }
  }
  if (Array.isArray(extras)) {
    for (const ex of extras) {
      try {
        const full = path.resolve(path.dirname(p), ex.path);
        await fs.promises.mkdir(path.dirname(full), { recursive: true });
        await fs.promises.writeFile(full, Buffer.from(ex.data));
      } catch (err) { console.warn('Falha ao gravar arquivo auxiliar', ex.path, err); }
    }
  }
  // imagens e outros arquivos salvos à parte não entram na lista de recentes
  if (SUPPORTED.includes(path.extname(p).toLowerCase())) addRecent(p);
  return { ok: true };
});

// ---------------------------------------------------------------- notas do Obsidian
// [[Nota]] aponta para um arquivo em qualquer pasta do cofre

const vaultCache = new Map();

function vaultRootOf(dir) {
  for (const marker of ['.obsidian', '.git']) {
    let d = dir;
    for (let i = 0; i < 20; i++) {
      try { if (fs.statSync(path.join(d, marker)).isDirectory()) return d; } catch { /* segue subindo */ }
      const up = path.dirname(d);
      if (up === d) break;
      d = up;
    }
  }
  return dir;
}

function listVault(root) {
  const cached = vaultCache.get(root);
  if (cached && Date.now() - cached.time < 30000) return cached.files;
  const files = [];
  const stack = [root];
  while (stack.length && files.length < 60000) {
    const d = stack.pop();
    let ents;
    try { ents = fs.readdirSync(d, { withFileTypes: true }); } catch { continue; }
    for (const e of ents) {
      if (e.name.startsWith('.') || e.name === 'node_modules') continue;
      const p = path.join(d, e.name);
      if (e.isDirectory()) stack.push(p);
      else files.push(p);
    }
  }
  vaultCache.set(root, { time: Date.now(), files });
  return files;
}

ipcMain.handle('file:find-note', (_e, from, target) => {
  if (!from) return null;
  const name = String(target || '').split('#')[0].split('^')[0].trim();
  if (!name) return from;
  const dir = path.dirname(from);
  const rel = name.replace(/[\\/]+/g, path.sep);
  // nomes de nota podem ter ponto, como "Backend Node.js"
  const candidates = /\.md$/i.test(rel) ? [rel] : [rel + '.md', rel];
  const root = vaultRootOf(dir);
  for (const base of [dir, root]) {
    for (const c of candidates) {
      const p = path.resolve(base, c);
      try { if (fs.statSync(p).isFile()) return p; } catch { /* tenta o próximo */ }
    }
  }
  const tails = candidates.map((c) => path.sep + c.toLowerCase());
  const bases = candidates.map((c) => path.basename(c).toLowerCase());
  let best = null;
  for (const f of listVault(root)) {
    const lf = f.toLowerCase();
    if (tails.some((t) => lf.endsWith(t))) return f;
    if (!best && bases.includes(path.basename(lf))) best = f;
  }
  return best;
});

ipcMain.handle('file:open-dialog', async (event, opts = {}) => {
  const win = windowFor(event);
  const res = await dialog.showOpenDialog(win, {
    title: opts.title || 'Abrir',
    defaultPath: opts.defaultPath,
    properties: ['openFile', 'multiSelections'],
    filters: opts.filters || [
      { name: 'Todos os documentos', extensions: ['docx', 'docm', 'dotx', 'odt', 'md', 'markdown', 'txt', 'html', 'htm', 'rtf'] },
      { name: 'Documento do Word', extensions: ['docx', 'docm', 'dotx'] },
      { name: 'Texto OpenDocument', extensions: ['odt'] },
      { name: 'Markdown', extensions: ['md', 'markdown'] },
      { name: 'Página da Web', extensions: ['html', 'htm'] },
      { name: 'Texto sem formatação', extensions: ['txt', 'text', 'log', 'csv'] },
      { name: 'Rich Text', extensions: ['rtf'] },
      { name: 'Todos os arquivos', extensions: ['*'] },
    ],
  });
  return res.canceled ? [] : res.filePaths;
});

ipcMain.handle('file:image-dialog', async (event) => {
  const win = windowFor(event);
  const res = await dialog.showOpenDialog(win, {
    title: 'Inserir Imagem',
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Imagens', extensions: ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg', 'emf', 'wmf', 'tif', 'tiff', 'ico'] },
      { name: 'Todos os arquivos', extensions: ['*'] },
    ],
  });
  if (res.canceled) return [];
  const out = [];
  for (const p of res.filePaths) {
    const data = await fs.promises.readFile(p);
    out.push({ path: p, name: path.basename(p), data: new Uint8Array(data.buffer, data.byteOffset, data.byteLength) });
  }
  return out;
});

ipcMain.handle('file:save-dialog', async (event, opts = {}) => {
  const win = windowFor(event);
  const res = await dialog.showSaveDialog(win, {
    title: opts.title || 'Salvar como',
    defaultPath: opts.defaultPath,
    filters: opts.filters,
  });
  return res.canceled ? null : res.filePath;
});

ipcMain.handle('file:pick-folder', async (event) => {
  const win = windowFor(event);
  const res = await dialog.showOpenDialog(win, { properties: ['openDirectory'] });
  return res.canceled ? null : res.filePaths[0];
});

ipcMain.on('app:set-doc-state', (event, st) => {
  const win = windowFor(event);
  if (!win) return;
  if ('filePath' in st) win.docState.filePath = st.filePath;
  if ('dirty' in st) win.docState.dirty = !!st.dirty;
  if (st.title) win.setTitle(st.title);
  if (st.filePath && st.addRecent) addRecent(st.filePath);
});

ipcMain.handle('app:recent', () => getRecent());
ipcMain.handle('app:recent-add', (_e, p) => addRecent(p));
ipcMain.handle('app:recent-remove', (_e, p) => removeRecent(p));

ipcMain.handle('app:new-window', (_e, p) => {
  if (p) openFileInNewWindow(p);
  else createWindow(null);
  return true;
});

ipcMain.on('app:close-confirmed', (event) => {
  const win = windowFor(event);
  if (!win) return;
  win.docState.allowClose = true;
  win.close();
});

ipcMain.on('app:close-window', (event) => {
  const win = windowFor(event);
  if (win) win.close();
});

ipcMain.on('app:quit', () => app.quit());

ipcMain.on('app:set-overlay', (event, colors) => {
  const win = windowFor(event);
  if (!win || !win.setTitleBarOverlay) return;
  try { win.setTitleBarOverlay({ color: colors.color || TITLE_BAR_COLOR, symbolColor: colors.symbolColor || '#ffffff', height: 40 }); } catch { /* ignora */ }
});

ipcMain.handle('app:save-prefs', (_e, prefs) => {
  const cur = readJson(prefsFile, {});
  const next = { ...cur, ...prefs };
  writeJson(prefsFile, next);
  return next;
});

ipcMain.handle('app:print', (event) => {
  const win = windowFor(event);
  return new Promise((resolve) => {
    win.webContents.print({ silent: false, printBackground: true, margins: { marginType: 'none' } }, (ok, reason) => resolve({ ok, reason }));
  });
});

ipcMain.handle('app:export-pdf', async (event, p, opts = {}) => {
  const win = windowFor(event);
  try {
    const data = await win.webContents.printToPDF({
      printBackground: true,
      preferCSSPageSize: true,
      margins: { marginType: 'none' },
      pageSize: opts.pageSize || 'A4',
      landscape: !!opts.landscape,
      generateDocumentOutline: true,
      generateTaggedPDF: true,
    });
    await fs.promises.writeFile(p, data);
    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
});

ipcMain.handle('shell:open-external', (_e, url) => {
  if (/^(https?|mailto):/i.test(url)) return shell.openExternal(url);
  return false;
});

ipcMain.handle('shell:open-path', (_e, p) => shell.openPath(p));
ipcMain.handle('shell:show-item', (_e, p) => { shell.showItemInFolder(p); return true; });

ipcMain.on('edit:action', (event, action) => {
  const allowed = ['cut', 'copy', 'paste', 'pasteAndMatchStyle', 'selectAll', 'delete'];
  if (allowed.includes(action)) event.sender[action]();
});

ipcMain.on('spell:replace', (event, word) => { event.sender.replaceMisspelling(word); });
ipcMain.on('spell:add', (_event, word) => { session.defaultSession.addWordToSpellCheckerDictionary(word); });
ipcMain.handle('spell:languages', () => ({
  available: session.defaultSession.availableSpellCheckerLanguages || [],
  current: session.defaultSession.getSpellCheckerLanguages(),
}));
ipcMain.handle('spell:set-languages', (_e, langs) => {
  try {
    session.defaultSession.setSpellCheckerLanguages(langs);
    const cur = readJson(prefsFile, {});
    writeJson(prefsFile, { ...cur, spellLanguages: langs });
    return true;
  } catch (err) { return false; }
});

ipcMain.handle('fonts:list', () => listSystemFonts());

ipcMain.handle('image:convert', (_e, data, ext, width, height) => convertImage(data, ext, width, height));

// ---------------------------------------------------------------- modo de teste local
// Ativado só com a variável REDATOR_TEST apontando para um script de teste.
// O script roda na página e conversa com este processo por mensagens de console.

function attachTestHarness(win) {
  const wc = win.webContents;
  const outDir = process.env.REDATOR_TEST_OUT || os.tmpdir();
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  wc.on('did-finish-load', async () => {
    await sleep(Number(process.env.REDATOR_TEST_DELAY || 1500));
    try {
      const script = fs.readFileSync(process.env.REDATOR_TEST, 'utf8');
      await wc.executeJavaScript(script, true);
    } catch (err) {
      console.log('TEST-ERR', err && err.message);
    }
  });
  wc.on('console-message', async (e) => {
    const msg = (e && e.message) || '';
    const level = e && e.level === 'error' ? 3 : 0;
    if (!msg.startsWith('TEST:')) { if (msg.startsWith('SELFTEST') || msg.startsWith('LOG')) console.log(msg); if (level >= 3 || (e && e.level === 'error')) console.log('CONSOLE-ERR', msg); return; }
    const [, cmd, ...rest] = msg.split(':');
    const arg = rest.join(':');
    try {
      if (cmd === 'SHOT') {
        const img = await wc.capturePage();
        fs.writeFileSync(path.join(outDir, arg), img.toPNG());
        console.log('TEST shot', arg);
      } else if (cmd === 'KEYS') {
        for (const ev of JSON.parse(arg)) {
          if (ev.wait) { await sleep(ev.wait); continue; }
          if (ev.text) {
            for (const ch of ev.text) { wc.sendInputEvent({ type: 'char', keyCode: ch }); await sleep(8); }
            continue;
          }
          const mods = ev.mods || [];
          wc.sendInputEvent({ type: 'keyDown', keyCode: ev.key, modifiers: mods });
          if (ev.char) wc.sendInputEvent({ type: 'char', keyCode: ev.char, modifiers: mods });
          wc.sendInputEvent({ type: 'keyUp', keyCode: ev.key, modifiers: mods });
          await sleep(20);
        }
        wc.executeJavaScript('window.__keysDone = (window.__keysDone || 0) + 1', true);
      } else if (cmd === 'MOUSE') {
        const ev = JSON.parse(arg);
        const base = { x: Math.round(ev.x), y: Math.round(ev.y), button: ev.button || 'left', clickCount: ev.count || 1 };
        wc.sendInputEvent({ type: 'mouseMove', x: base.x, y: base.y });
        wc.sendInputEvent({ ...base, type: 'mouseDown' });
        wc.sendInputEvent({ ...base, type: 'mouseUp' });
        wc.executeJavaScript('window.__mouseDone = (window.__mouseDone || 0) + 1', true);
      } else if (cmd === 'PDFSHOT') {
        const [pdfPath, prefix, zoom] = arg.split('|');
        const pv = new BrowserWindow({ width: 1000, height: 1300, show: true, webPreferences: { plugins: true } });
        await pv.loadURL('file:///' + pdfPath.replace(/\\/g, '/') + '#toolbar=0&zoom=' + (zoom || 40));
        await sleep(3000);
        const img = await pv.webContents.capturePage();
        fs.writeFileSync(path.join(outDir, prefix + '.png'), img.toPNG());
        pv.destroy();
        console.log('TEST pdfshot', prefix);
      } else if (cmd === 'OSCLICK') {
        // clique real do sistema operacional, para testar o menu de ortografia
        const ev = JSON.parse(arg);
        const b = win.getContentBounds();
        const disp = screen.getDisplayMatching(b);
        const sf = disp.scaleFactor || 1;
        const px = Math.round((b.x + ev.x) * sf);
        const py = Math.round((b.y + ev.y) * sf);
        const down = ev.button === 'right' ? '0x0008' : '0x0002';
        const up = ev.button === 'right' ? '0x0010' : '0x0004';
        const ps = "Add-Type -Namespace W -Name U -MemberDefinition '[DllImport(\"user32.dll\")] public static extern bool SetProcessDPIAware(); [DllImport(\"user32.dll\")] public static extern bool SetCursorPos(int x,int y); [DllImport(\"user32.dll\")] public static extern void mouse_event(int f,int x,int y,int d,int e);'; [W.U]::SetProcessDPIAware() | Out-Null; [W.U]::SetCursorPos(" + px + ',' + py + '); Start-Sleep -Milliseconds 120; [W.U]::mouse_event(' + down + ',0,0,0,0); Start-Sleep -Milliseconds 60; [W.U]::mouse_event(' + up + ',0,0,0,0)';
        win.focus();
        await new Promise((r) => execFile('powershell.exe', ['-NoProfile', '-Command', ps], { windowsHide: true }, () => r()));
        wc.executeJavaScript('window.__mouseDone = (window.__mouseDone || 0) + 1', true);
      } else if (cmd === 'CLIP') {
        const { clipboard, ClipboardItem } = require('electron');
        const data = JSON.parse(arg);
        await clipboard.write([new ClipboardItem({ 'text/html': data.html, 'text/plain': data.text || '' })]);
        wc.executeJavaScript('window.__clipDone = (window.__clipDone || 0) + 1', true);
      } else if (cmd === 'LOG') {
        console.log('TEST', arg);
      } else if (cmd === 'QUIT') {
        win.docState.allowClose = true;
        app.exit(0);
      }
    } catch (err) {
      console.log('TEST-ERR', cmd, err && err.message);
    }
  });
}

// ---------------------------------------------------------------- sugestões do corretor do Windows
// O Chromium sublinha com o corretor nativo do Windows; aqui pedimos as sugestões
// ao mesmo corretor, por COM, num processo auxiliar que fica aberto.

const SPELL_CS = `
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
[ComImport, Guid("B6FD0B71-E2BC-4653-8D05-F197E412770B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IRSpellChecker {
  [return: MarshalAs(UnmanagedType.LPWStr)] string GetLanguageTag();
  [return: MarshalAs(UnmanagedType.IUnknown)] object Check([MarshalAs(UnmanagedType.LPWStr)] string text);
  [PreserveSig] int Suggest([MarshalAs(UnmanagedType.LPWStr)] string word, out IEnumString value);
  void Add([MarshalAs(UnmanagedType.LPWStr)] string word);
  void Ignore([MarshalAs(UnmanagedType.LPWStr)] string word);
}
[ComImport, Guid("8E018A9D-2415-4677-BF08-794EA61F94BB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IRSpellCheckerFactory {
  [return: MarshalAs(UnmanagedType.IUnknown)] object GetSupportedLanguages();
  int IsSupported([MarshalAs(UnmanagedType.LPWStr)] string languageTag);
  IRSpellChecker CreateSpellChecker([MarshalAs(UnmanagedType.LPWStr)] string languageTag);
}
[ComImport, Guid("7AB36653-1796-484B-BDFA-E74F1DB7C1DC")]
public class RSpellCheckerFactory {}
public static class RSpell {
  static IRSpellChecker checker;
  public static string Init(string lang) {
    var f = (IRSpellCheckerFactory)new RSpellCheckerFactory();
    if (f.IsSupported(lang) == 0) return "unsupported";
    checker = f.CreateSpellChecker(lang);
    return "ok";
  }
  public static string Query(string word) {
    IEnumString e;
    int hr = checker.Suggest(word, out e);
    var list = new List<string>();
    if (e != null) {
      var buf = new string[1];
      while (list.Count < 8 && e.Next(1, buf, IntPtr.Zero) == 0) list.Add(buf[0]);
    }
    return (hr == 0 ? "1" : "0") + "\\t" + string.Join("\\t", list);
  }
  public static void Add(string word) { checker.Add(word); }
}`;

let spellProc = null;
let spellQueue = [];
let spellBuf = '';
let spellLang = 'pt-BR';

function startSpellHelper() {
  if (spellProc || process.platform !== 'win32') return;
  const script = [
    '[Console]::InputEncoding = [Text.Encoding]::UTF8',
    '[Console]::OutputEncoding = [Text.Encoding]::UTF8',
    'Add-Type -TypeDefinition @"\n' + SPELL_CS + '\n"@',
    '$r = [RSpell]::Init("' + spellLang + '")',
    '[Console]::Out.WriteLine("init`t" + $r); [Console]::Out.Flush()',
    'while ($true) { $l = [Console]::In.ReadLine(); if ($l -eq $null) { break }',
    '  try { if ($l.StartsWith("+")) { [RSpell]::Add($l.Substring(1)); $o = "added" } else { $o = [RSpell]::Query($l) } } catch { $o = "err`t" + $_.Exception.Message }',
    '  [Console]::Out.WriteLine($o); [Console]::Out.Flush() }',
  ].join('\n');
  try {
    const encoded = Buffer.from(script, 'utf16le').toString('base64');
    spellProc = require('child_process').spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded], { windowsHide: true, stdio: ['pipe', 'pipe', 'ignore'] });
  } catch { spellProc = null; return; }
  spellProc.stdout.setEncoding('utf8');
  spellProc.stdout.on('data', (chunk) => {
    spellBuf += chunk;
    let i;
    while ((i = spellBuf.indexOf('\n')) >= 0) {
      const line = spellBuf.slice(0, i).replace(/\r$/, '');
      spellBuf = spellBuf.slice(i + 1);
      if (line.startsWith('init\t')) continue;
      const job = spellQueue.shift();
      if (job) job(line);
    }
  });
  spellProc.on('exit', () => { spellProc = null; for (const j of spellQueue) j(''); spellQueue = []; });
}

function spellQuery(line) {
  startSpellHelper();
  if (!spellProc) return Promise.resolve('');
  return new Promise((resolve) => {
    const t = setTimeout(() => resolve(''), 8000);
    spellQueue.push((out) => { clearTimeout(t); resolve(out); });
    spellProc.stdin.write(line.replace(/[\r\n]/g, ' ') + '\n');
  });
}

ipcMain.handle('spell:suggest', async (_e, word) => {
  const out = await spellQuery(String(word || ''));
  const parts = out.split('\t');
  if (parts[0] !== '1' && parts[0] !== '0') return { misspelled: false, suggestions: [] };
  return { misspelled: parts[0] === '1', suggestions: parts.slice(1).filter(Boolean) };
});

ipcMain.handle('spell:add-word', async (_e, word) => {
  await spellQuery('+' + String(word || ''));
  try { session.defaultSession.addWordToSpellCheckerDictionary(word); } catch { /* ignora */ }
  return true;
});

app.on('will-quit', () => { if (spellProc) { try { spellProc.kill(); } catch { /* ignora */ } } });

// ---------------------------------------------------------------- fontes do sistema

let fontCache = null;
function listSystemFonts() {
  if (fontCache) return Promise.resolve(fontCache);
  const keys = ['HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts', 'HKCU\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts'];
  const run = (key) => new Promise((resolve) => {
    execFile('reg', ['query', key], { windowsHide: true, maxBuffer: 8 * 1024 * 1024 }, (err, stdout) => resolve(err ? '' : stdout));
  });
  return Promise.all(keys.map(run)).then((outs) => {
    const names = new Set();
    const styleWords = /\b(Bold|Italic|Oblique|Light|Semilight|SemiLight|Semibold|SemiBold|Demibold|Black|Heavy|Medium|Thin|ExtraLight|Extra Light|ExtraBold|Extra Bold|UltraLight|Condensed|Narrow|Regular|Book|Negreta|Cursiva)\b/g;
    for (const out of outs) {
      for (const line of out.split(/\r?\n/)) {
        const m = line.match(/^\s{4}(.+?)\s+REG_SZ\s+/);
        if (!m) continue;
        let name = m[1].replace(/\s*\((TrueType|OpenType|All res|VGA res|8514\/a res)[^)]*\)\s*$/i, '');
        for (const part of name.split('&')) {
          let n = part.trim().replace(styleWords, '').replace(/\s{2,}/g, ' ').trim();
          if (n && !/^\d/.test(n) && n.length < 60) names.add(n);
        }
      }
    }
    fontCache = [...names].sort((a, b) => a.localeCompare(b, 'pt-BR'));
    return fontCache;
  });
}

// ---------------------------------------------------------------- conversão de imagens (EMF, WMF, TIFF)

function convertImage(data, ext, width, height) {
  if (process.platform !== 'win32') return Promise.resolve(null);
  return new Promise((resolve) => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'redator-img-'));
    const src = path.join(dir, 'origem.' + String(ext || 'emf').replace(/[^a-z0-9]/gi, ''));
    const dst = path.join(dir, 'saida.png');
    fs.writeFileSync(src, Buffer.from(data));
    const w = Math.max(1, Math.min(4000, Math.round(width || 0)));
    const h = Math.max(1, Math.min(4000, Math.round(height || 0)));
    const script = [
      'Add-Type -AssemblyName System.Drawing',
      `$img = [System.Drawing.Image]::FromFile('${src.replace(/'/g, "''")}')`,
      `$w = ${width ? w : '$img.Width'}; $h = ${height ? h : '$img.Height'}`,
      'if ($w -lt 1) { $w = $img.Width }; if ($h -lt 1) { $h = $img.Height }',
      '$bmp = New-Object System.Drawing.Bitmap($w, $h)',
      '$g = [System.Drawing.Graphics]::FromImage($bmp)',
      '$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality',
      '$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic',
      '$g.Clear([System.Drawing.Color]::Transparent)',
      '$g.DrawImage($img, 0, 0, $w, $h)',
      `$bmp.Save('${dst.replace(/'/g, "''")}', [System.Drawing.Imaging.ImageFormat]::Png)`,
      '$g.Dispose(); $bmp.Dispose(); $img.Dispose()',
    ].join('; ');
    execFile('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', script], { windowsHide: true, timeout: 20000 }, (err) => {
      let out = null;
      try { if (!err) { const b = fs.readFileSync(dst); out = new Uint8Array(b.buffer, b.byteOffset, b.byteLength); } } catch { out = null; }
      try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* ignora */ }
      resolve(out);
    });
  });
}
