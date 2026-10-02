'use strict';

const { app, BrowserWindow, ipcMain, dialog, shell, Menu, session, screen, nativeImage } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { execFile } = require('child_process');
const { createFileStore, hashBytes } = require('./services/file-store.cjs');
const { createTrustedIpc } = require('./services/trusted-ipc.cjs');
const { createLineHelper } = require('./services/line-helper.cjs');
const { normalizePreferences } = require('./services/preferences.cjs');
const { encodeSnapshot, readRecovery, scanRecovery: findRecovery, recoveryFiles, RECOVER_ID } = require('./services/recovery-snapshot.cjs');
const fileStore = createFileStore();

const APP_NAME = 'Redator';
const TITLE_BAR_COLOR = '#185abd';
const TEXT_EXTS = ['txt', 'text', 'log', 'json', 'jsonc', 'json5', 'xml', 'xsd', 'xsl', 'svg', 'yaml', 'yml', 'toml', 'ini', 'cfg', 'conf', 'config', 'properties', 'env',
  'sql', 'ps1', 'psm1', 'psd1', 'bat', 'cmd', 'sh', 'bash', 'zsh', 'py', 'js', 'mjs', 'cjs', 'ts', 'tsx', 'jsx', 'vue', 'css', 'scss', 'sass', 'less',
  'java', 'kt', 'kts', 'groovy', 'gradle', 'c', 'h', 'cpp', 'cc', 'hpp', 'cs', 'go', 'rs', 'php', 'rb', 'swift', 'vb', 'vbs', 'lua', 'r', 'pl', 'dart', 'scala',
  'srt', 'vtt', 'sub', 'rst', 'adoc', 'asciidoc', 'org', 'tex', 'bib', 'diff', 'patch', 'nfo', 'reg', 'tsv', 'gitignore', 'gitattributes', 'editorconfig', 'dockerfile'];
const SUPPORTED = ['.docx', '.docm', '.dotx', '.dotm', '.doc', '.dot', '.md', '.markdown', '.csv', '.tsv', '.html', '.htm', '.xhtml', '.odt', '.ott', '.fodt', '.rtf', '.pdf', '.epub', '.mht', '.mhtml', '.eml', ...TEXT_EXTS.map((e) => '.' + e)];

app.setName(APP_NAME);
if (process.platform === 'win32') app.setAppUserModelId('br.redator.app');
// testes rodam com a janela escondida, sem aparecer na tela nem roubar o teclado;
// REDATOR_TEST_SHOW=1 mostra a janela quando é preciso ver o que acontece
const HIDDEN_TEST = !!process.env.REDATOR_TEST && !process.env.REDATOR_TEST_SHOW;
if (process.env.REDATOR_TEST) {
  app.setPath('userData', path.join(process.env.REDATOR_TEST_OUT || os.tmpdir(), 'redator-test-profile'));
  // máquinas de CI nem sempre têm placa de vídeo
  app.disableHardwareAcceleration();
}

const userData = app.getPath('userData');
const recentFile = path.join(userData, 'recentes.json');
const windowStateFile = path.join(userData, 'janela.json');
const prefsFile = path.join(userData, 'preferencias.json');
const recoverDir = path.join(userData, 'Recuperacao');

// ---------------------------------------------------------------- utilidades

function readJson(file, fallback) {
  try {
    const value = JSON.parse(fs.readFileSync(file, 'utf8'));
    if (!value || typeof value !== 'object' || Array.isArray(value) !== Array.isArray(fallback)) return fallback;
    return value;
  } catch { return fallback; }
}

function writeJson(file, value) {
  const tmp = file + '.' + require('crypto').randomUUID() + '.tmp';
  let fd;
  try {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fd = fs.openSync(tmp, 'wx');
    fs.writeFileSync(fd, JSON.stringify(value, null, 2), 'utf8');
    fs.fsyncSync(fd);
    fs.closeSync(fd);
    fd = undefined;
    fs.renameSync(tmp, file);
    return true;
  } catch (err) {
    console.error('Falha ao gravar', file, err);
    return false;
  } finally {
    if (fd !== undefined) { try { fs.closeSync(fd); } catch { /* encerrado */ } }
    try { fs.unlinkSync(tmp); } catch { /* publicado ou não criado */ }
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
  return list.filter((r) => r && typeof r.path === 'string' && r.path).slice(0, 40).map((r) => ({
    path: r.path, name: path.basename(r.path), folder: path.dirname(r.path), time: Number.isFinite(r.time) ? r.time : 0,
  }));
}

function addRecent(p) {
  if (typeof p !== 'string' || !p) return getRecent();
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

// ---------------------------------------------------------------- AutoRecuperação
// Cada janela com alterações guarda uma cópia do documento aqui. A cópia some
// quando a janela fecha normalmente; se o programa ou o PC cair, ela fica e é
// oferecida na próxima abertura.

let leftovers = [];
const recoveryCleanup = new Set();
let waitingForCleanup = false;

app.on('will-quit', (event) => {
  if (!recoveryCleanup.size) return;
  event.preventDefault();
  if (waitingForCleanup) return;
  waitingForCleanup = true;
  (async () => {
    while (recoveryCleanup.size) await Promise.allSettled([...recoveryCleanup]);
    waitingForCleanup = false;
    app.quit();
  })();
});

function recoverFiles(id) {
  return recoveryFiles(recoverDir, id);
}

function removeRecovery(id) {
  if (typeof id !== 'string' || !RECOVER_ID.test(id)) return Promise.resolve();
  const f = recoverFiles(id);
  const paths = [f.snapshot, f.data, f.meta, f.data + '.tmp', f.meta + '.tmp'];
  const job = fileStore.enqueue(paths, async () => {
    for (const p of paths) {
      try { await fs.promises.unlink(p); }
      catch (err) { if (err.code !== 'ENOENT') console.warn('Falha ao apagar a recuperação:', err.message); }
    }
  });
  recoveryCleanup.add(job);
  job.then(() => recoveryCleanup.delete(job), () => recoveryCleanup.delete(job));
  return job;
}

function scanRecovery() {
  return findRecovery(recoverDir);
}

// ---------------------------------------------------------------- janelas

const windows = new Set();
const trustedIpc = createTrustedIpc(ipcMain, {
  windows,
  getWindow: windowFor,
  appUrl: require('url').pathToFileURL(path.join(__dirname, 'app', 'index.html')).href,
});

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

function createWindow(filePath, opts = {}) {
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
      // nos testes a janela fica escondida e não pode ter os relógios freados
      backgroundThrottling: !HIDDEN_TEST,
    },
  });

  win.docState = { filePath: filePath || null, dirty: false, allowClose: false, recover: opts.recover || null, crashes: 0 };
  win.recoverId = null;
  win.recoveryEpoch = 0;
  win.recoveryClaim = null;
  windows.add(win);

  win.loadFile(path.join(__dirname, 'app', 'index.html'));

  win.once('ready-to-show', () => {
    if (HIDDEN_TEST) { win.webContents.focus(); return; }
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

  win.on('closed', async () => {
    windows.delete(win);
    // fechou pelo caminho normal: o usuário salvou ou escolheu não salvar
    const decided = !win.docState.dirty || win.docState.allowClose;
    if (decided && !win.keepRecovery) {
      win.recoveryEpoch++;
      if (win.recoverId && !win.recoveryClaim) await removeRecovery(win.recoverId);
    }
  });
  // desligamento do Windows: a cópia fica para a próxima abertura
  win.on('session-end', () => { win.keepRecovery = true; });

  win.on('maximize', () => win.webContents.send('app:window-state', { maximized: true }));
  win.on('unmaximize', () => win.webContents.send('app:window-state', { maximized: false }));
  win.on('focus', () => win.webContents.send('app:focus'));

  const wc = win.webContents;
  wc.on('will-navigate', (e) => e.preventDefault());
  wc.on('render-process-gone', (_e, details) => {
    if (details.reason === 'clean-exit' || win.isDestroyed()) return;
    console.warn('A janela parou de responder:', details.reason);
    const st = win.docState;
    st.dirty = false;
    // se a mesma falha se repetir, deixa a janela fechar em vez de recarregar sem fim
    if (++st.crashes > 2) { st.allowClose = true; win.keepRecovery = true; return; }
    if (win.recoverId) {
      const f = recoverFiles(win.recoverId);
      if (fs.existsSync(f.snapshot) || fs.existsSync(f.meta)) st.recover = win.recoverId;
    }
    setTimeout(() => { if (!win.isDestroyed()) wc.reload(); }, 400);
  });
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
    leftovers = scanRecovery();
    const files = filesFromArgv(process.argv, process.cwd());
    if (files.length) files.forEach((f) => createWindow(f));
    else createWindow(null);
  });
}

app.on('window-all-closed', () => { app.quit(); });

// ---------------------------------------------------------------- ortografia

function configureSpellchecker() {
  const ses = session.defaultSession;
  const prefs = normalizePreferences(readJson(prefsFile, {}));
  const wanted = prefs.spellLanguages || ['pt-BR'];
  try {
    const available = ses.availableSpellCheckerLanguages || [];
    const langs = wanted.filter((l) => !available.length || available.includes(l));
    ses.setSpellCheckerLanguages(langs.length ? langs : wanted);
    spellLang = langs[0] || wanted[0] || 'pt-BR';
  } catch (err) {
    console.warn('Idiomas de ortografia indisponíveis', err.message);
  }
  ses.setPermissionRequestHandler((_wc, permission, cb) => {
    cb(permission === 'local-fonts' || permission === 'clipboard-read' || permission === 'clipboard-sanitized-write');
  });
  ses.setPermissionCheckHandler((_wc, permission) => permission === 'local-fonts' || permission === 'clipboard-read' || permission === 'clipboard-sanitized-write');
}

// ---------------------------------------------------------------- IPC

trustedIpc.handle('app:init', (event) => {
  const win = windowFor(event);
  const recover = win ? win.docState.recover : null;
  if (win) win.docState.recover = null;
  // o que sobrou da sessão anterior é oferecido uma vez, na primeira janela
  const recovered = leftovers.filter((e) => ![...windows].some((w) => w.recoverId === e.id));
  leftovers = [];
  return {
    filePath: win ? win.docState.filePath : null,
    recover,
    recovered,
    recent: getRecent(),
    version: app.getVersion(),
    userName: os.userInfo().username,
    prefs: normalizePreferences(readJson(prefsFile, {})),
    platform: process.platform,
  };
});

trustedIpc.handle('recover:save', async (event, data, meta) => {
  const win = windowFor(event);
  if (!win || win.recoveryClaim) return false;
  if (!win.recoverId) win.recoverId = Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
  const f = recoverFiles(win.recoverId);
  const id = win.recoverId;
  const epoch = win.recoveryEpoch;
  try {
    await fs.promises.mkdir(recoverDir, { recursive: true });
    if (win.recoveryEpoch !== epoch || win.recoverId !== id) return null;
    const packet = encodeSnapshot(data, meta);
    const result = await fileStore.writeFile(f.snapshot, packet);
    if (!result.ok) throw new Error(result.error);
    if (win.recoveryEpoch !== epoch || win.recoverId !== id) {
      if (!win.recoveryClaim || win.recoveryClaim.previous !== id) await removeRecovery(id);
      return null;
    }
    return id;
  } catch (err) {
    console.warn('Falha na AutoRecuperação', err.message);
    return null;
  }
});

trustedIpc.handle('recover:clear', async (event) => {
  const win = windowFor(event);
  if (win && !win.recoveryClaim) {
    win.recoveryEpoch++;
    if (win.recoverId) await removeRecovery(win.recoverId);
  }
  return true;
});

// a janela assume a cópia: as próximas gravações vão para ela e fechar normalmente a apaga
trustedIpc.handle('recover:take', async (event, id) => {
  const win = windowFor(event);
  if (!win || win.recoveryClaim || typeof id !== 'string' || !RECOVER_ID.test(id)) return null;
  if ([...windows].some((w) => w !== win && w.recoverId === id)) return null;
  const previous = win.recoverId;
  // Claim before disk I/O so two windows cannot take the same copy.
  win.recoveryClaim = { id, previous };
  win.recoverId = id;
  try {
    const { meta, data } = await readRecovery(recoverDir, id);
    return { meta, data: new Uint8Array(data.buffer, data.byteOffset, data.byteLength) };
  } catch {
    if (win.recoverId === id) win.recoverId = previous;
    win.recoveryClaim = null;
    return null;
  }
});

// Adoption completes only after the renderer parsed the document successfully.
// Until then, closing the window preserves the original recovery copy.
trustedIpc.handle('recover:confirm', async (event, id) => {
  const win = windowFor(event);
  if (!win || !win.recoveryClaim || win.recoveryClaim.id !== id) return false;
  const { previous } = win.recoveryClaim;
  win.recoveryClaim = null;
  if (previous && previous !== id) await removeRecovery(previous);
  return true;
});

trustedIpc.handle('recover:release', (event, id) => {
  const win = windowFor(event);
  if (!win || !win.recoveryClaim || win.recoveryClaim.id !== id) return false;
  win.recoverId = win.recoveryClaim.previous;
  win.recoveryClaim = null;
  return true;
});

trustedIpc.handle('recover:open', (_e, id) => {
  if (typeof id !== 'string' || !RECOVER_ID.test(id)) return false;
  const opened = [...windows].find((w) => w.recoverId === id || w.docState.recover === id);
  if (opened) { if (!HIDDEN_TEST) opened.focus(); return true; }
  createWindow(null, { recover: id });
  return true;
});

trustedIpc.handle('recover:discard', async (_e, ids) => {
  for (const id of Array.isArray(ids) ? ids : []) {
    if (![...windows].some((w) => w.recoverId === id || w.docState.recover === id)) await removeRecovery(id);
  }
  return true;
});

trustedIpc.handle('file:read', async (_event, p) => {
  const data = await fs.promises.readFile(p);
  let readonly = false;
  try { await fs.promises.access(p, fs.constants.W_OK); } catch { readonly = true; }
  const st = await fs.promises.stat(p);
  return { data: new Uint8Array(data.buffer, data.byteOffset, data.byteLength), readonly, mtime: st.mtimeMs, size: st.size, hash: hashBytes(data), name: path.basename(p), dir: path.dirname(p) };
});

trustedIpc.handle('file:stat', async (_event, p) => {
  try { const st = await fs.promises.stat(p); return st.isFile() ? { mtime: st.mtimeMs, size: st.size } : null; } catch { return null; }
});

trustedIpc.handle('file:exists', (_event, p) => {
  try { return fs.statSync(p).isFile(); } catch { return false; }
});

trustedIpc.handle('file:write', async (_event, p, data, extras, options) => {
  const result = await fileStore.writeFile(p, data, extras, options);
  if (!result.ok) return result;
  // imagens e outros arquivos salvos à parte não entram na lista de recentes
  if (SUPPORTED.includes(path.extname(p).toLowerCase())) addRecent(p);
  return result;
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

trustedIpc.handle('file:find-note', (_e, from, target) => {
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

trustedIpc.handle('file:open-dialog', async (event, opts = {}) => {
  const win = windowFor(event);
  const res = await dialog.showOpenDialog(win, {
    title: opts.title || 'Abrir',
    defaultPath: opts.defaultPath,
    properties: ['openFile', 'multiSelections'],
    filters: opts.filters || [
      { name: 'Todos os documentos', extensions: ['docx', 'docm', 'dotx', 'dotm', 'doc', 'dot', 'odt', 'ott', 'fodt', 'md', 'markdown', 'txt', 'csv', 'html', 'htm', 'xhtml', 'mht', 'mhtml', 'rtf', 'pdf', 'epub', 'eml'] },
      { name: 'Documento do Word', extensions: ['docx', 'docm', 'dotx'] },
      { name: 'Documento do Word 97-2003', extensions: ['doc', 'dot'] },
      { name: 'Texto OpenDocument', extensions: ['odt', 'ott', 'fodt'] },
      { name: 'Markdown', extensions: ['md', 'markdown'] },
      { name: 'Página da Web', extensions: ['html', 'htm'] },
      { name: 'Texto sem formatação', extensions: ['txt', 'text', 'log', 'csv'] },
      { name: 'Texto e código', extensions: TEXT_EXTS },
      { name: 'Rich Text', extensions: ['rtf'] },
      { name: 'PDF', extensions: ['pdf'] },
      { name: 'Livro digital EPUB', extensions: ['epub'] },
      { name: 'Página da Web em arquivo único', extensions: ['mht', 'mhtml'] },
      { name: 'Mensagem de e-mail', extensions: ['eml'] },
      { name: 'Todos os arquivos', extensions: ['*'] },
    ],
  });
  return res.canceled ? [] : res.filePaths;
});

trustedIpc.handle('file:image-dialog', async (event) => {
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

trustedIpc.handle('file:save-dialog', async (event, opts = {}) => {
  const win = windowFor(event);
  const res = await dialog.showSaveDialog(win, {
    title: opts.title || 'Salvar como',
    defaultPath: opts.defaultPath,
    filters: opts.filters,
  });
  return res.canceled ? null : res.filePath;
});

trustedIpc.handle('file:pick-folder', async (event) => {
  const win = windowFor(event);
  const res = await dialog.showOpenDialog(win, { properties: ['openDirectory'] });
  return res.canceled ? null : res.filePaths[0];
});

trustedIpc.on('app:set-doc-state', (event, st) => {
  const win = windowFor(event);
  if (!win || !st || typeof st !== 'object' || Array.isArray(st)) return;
  if (st.filePath === null || typeof st.filePath === 'string') win.docState.filePath = st.filePath;
  if (typeof st.dirty === 'boolean') win.docState.dirty = st.dirty;
  if (typeof st.title === 'string' && st.title) win.setTitle(st.title);
  if (typeof st.filePath === 'string' && st.filePath && st.addRecent === true) addRecent(st.filePath);
});

trustedIpc.handle('app:recent', () => getRecent());
trustedIpc.handle('app:recent-add', (_e, p) => addRecent(p));
trustedIpc.handle('app:recent-remove', (_e, p) => removeRecent(p));

trustedIpc.handle('app:new-window', (_e, p) => {
  if (p) openFileInNewWindow(p);
  else createWindow(null);
  return true;
});

trustedIpc.on('app:close-confirmed', (event) => {
  const win = windowFor(event);
  if (!win) return;
  win.docState.allowClose = true;
  win.close();
});

trustedIpc.on('app:close-window', (event) => {
  const win = windowFor(event);
  if (win) win.close();
});

trustedIpc.on('app:quit', () => app.quit());

trustedIpc.on('app:set-overlay', (event, colors) => {
  const win = windowFor(event);
  if (!win || !win.setTitleBarOverlay || !colors || typeof colors !== 'object') return;
  try { win.setTitleBarOverlay({ color: colors.color || TITLE_BAR_COLOR, symbolColor: colors.symbolColor || '#ffffff', height: 40 }); } catch { /* ignora */ }
});

trustedIpc.handle('app:save-prefs', (_e, prefs) => {
  const cur = normalizePreferences(readJson(prefsFile, {}));
  const next = { ...cur, ...normalizePreferences(prefs) };
  if (!writeJson(prefsFile, next)) throw new Error('Não foi possível guardar as preferências.');
  return next;
});

trustedIpc.handle('app:print', (event) => {
  const win = windowFor(event);
  return new Promise((resolve) => {
    win.webContents.print({ silent: false, printBackground: true, margins: { marginType: 'none' } }, (ok, reason) => resolve({ ok, reason }));
  });
});

trustedIpc.handle('app:export-pdf', async (event, p, opts = {}) => {
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
    return await fileStore.writeFile(p, data);
  } catch (err) {
    return { ok: false, error: err.message };
  }
});

trustedIpc.handle('shell:open-external', (_e, url) => {
  if (/^(https?|mailto):/i.test(url)) return shell.openExternal(url);
  return false;
});

trustedIpc.handle('shell:open-path', (_e, p) => shell.openPath(p));

// ---------------------------------------------------------------- aplicativo padrão
// O Windows não deixa um programa se declarar padrão: mostra para que tipos o
// Redator já é o escolhido e abre a página dele nas configurações.
const DEFAULT_EXTS = ['.docx', '.doc', '.pdf', '.md', '.txt', '.rtf', '.odt', '.epub', '.csv'];

function userChoice(ext) {
  const base = 'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\' + ext;
  const query = (key) => new Promise((resolve) => {
    execFile('reg', ['query', base + '\\' + key, '/v', 'ProgId'], { windowsHide: true }, (err, stdout) => {
      const m = !err && /ProgId\s+REG_SZ\s+(.+)/.exec(stdout || '');
      resolve(m ? m[1].trim() : null);
    });
  });
  // o Windows 11 mais novo guarda a escolha em UserChoiceLatest
  return query('UserChoiceLatest').then((v) => v || query('UserChoice'));
}

trustedIpc.handle('app:default-status', async () => {
  if (process.platform !== 'win32') return {};
  const out = {};
  await Promise.all(DEFAULT_EXTS.map(async (e) => {
    const prog = await userChoice(e);
    out[e] = !!prog && /redator/i.test(prog);
  }));
  return out;
});

// PDF convertido: o original abre no visualizador de PDF do Chromium, fiel ao arquivo
trustedIpc.handle('pdf:view-original', (event, p) => {
  if (!p || !/\.pdf$/i.test(p) || !isSupportedFile(p)) return false;
  const parent = windowFor(event);
  const pb = parent && !parent.isDestroyed() ? parent.getBounds() : { x: undefined, y: undefined, width: 1100, height: 860 };
  const iconPath = path.join(__dirname, 'app', 'icons', 'redator.ico');
  const w = new BrowserWindow({
    x: pb.x != null ? pb.x + 36 : undefined, y: pb.y != null ? pb.y + 36 : undefined,
    width: Math.max(800, pb.width - 72), height: Math.max(600, pb.height - 72),
    title: path.basename(p) + ' - original - Redator', backgroundColor: '#525659', show: !HIDDEN_TEST,
    icon: fs.existsSync(iconPath) ? iconPath : undefined,
    webPreferences: { plugins: true, contextIsolation: true, nodeIntegration: false, sandbox: true },
  });
  w.webContents.setWindowOpenHandler(({ url }) => {
    if (/^(https?|mailto):/i.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  w.webContents.on('will-navigate', (e, url) => {
    if (!/^file:/i.test(url)) { e.preventDefault(); if (/^(https?|mailto):/i.test(url)) shell.openExternal(url); }
  });
  w.webContents.on('page-title-updated', (e) => e.preventDefault());
  w.loadURL(require('url').pathToFileURL(p).href);
  return true;
});

trustedIpc.handle('app:open-default-apps', async () => {
  try { await shell.openExternal('ms-settings:defaultapps?registeredAppUser=Redator'); } catch {
    try { await shell.openExternal('ms-settings:defaultapps'); } catch { return false; }
  }
  return true;
});
trustedIpc.handle('shell:show-item', (_e, p) => { shell.showItemInFolder(p); return true; });

trustedIpc.on('edit:action', (event, action) => {
  const allowed = ['cut', 'copy', 'paste', 'pasteAndMatchStyle', 'selectAll', 'delete'];
  if (allowed.includes(action)) event.sender[action]();
});

trustedIpc.on('spell:replace', (event, word) => { event.sender.replaceMisspelling(word); });
trustedIpc.on('spell:add', (_event, word) => { session.defaultSession.addWordToSpellCheckerDictionary(word); });
trustedIpc.handle('spell:languages', () => ({
  available: session.defaultSession.availableSpellCheckerLanguages || [],
  current: session.defaultSession.getSpellCheckerLanguages(),
}));
trustedIpc.handle('spell:set-languages', (_e, langs) => {
  try {
    if (!Array.isArray(langs) || langs.some((lang) => typeof lang !== 'string' || !session.defaultSession.availableSpellCheckerLanguages.includes(lang))) return false;
    session.defaultSession.setSpellCheckerLanguages(langs);
    const cur = normalizePreferences(readJson(prefsFile, {}));
    writeJson(prefsFile, { ...cur, spellLanguages: langs });
    spellLang = langs[0] || 'pt-BR';
    spellHelper.stop();
    return true;
  } catch (err) { return false; }
});

trustedIpc.handle('fonts:list', () => listSystemFonts());

trustedIpc.handle('image:convert', (_e, data, ext, width, height) => convertImage(data, ext, width, height));

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
        // janela escondida não se redesenha sozinha: pede um quadro novo antes de fotografar
        wc.invalidate();
        await wc.executeJavaScript('new Promise((r) => { requestAnimationFrame(() => requestAnimationFrame(r)); setTimeout(r, 400); })', true).catch(() => {});
        // a primeira captura acorda a pintura e ainda devolve o quadro velho
        if (HIDDEN_TEST) { await wc.capturePage(undefined, { stayHidden: true }); await sleep(250); }
        const img = await wc.capturePage(undefined, { stayHidden: true });
        fs.writeFileSync(path.join(outDir, arg), img.toPNG());
        console.log('TEST shot', arg);
        await wc.executeJavaScript('window.__shotsDone = (window.__shotsDone || 0) + 1', true);
      } else if (cmd === 'RESIZE') {
        const size = JSON.parse(arg);
        win.setMinimumSize(320, 320);
        win.setContentSize(Math.max(320, Math.round(size.width)), Math.max(320, Math.round(size.height)));
        await wc.executeJavaScript('window.__resizeDone = (window.__resizeDone || 0) + 1', true);
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
        const pv = new BrowserWindow({ width: 1000, height: 1300, show: !HIDDEN_TEST, webPreferences: { plugins: true, backgroundThrottling: false } });
        await pv.loadURL('file:///' + pdfPath.replace(/\\/g, '/') + '#toolbar=0&zoom=' + (zoom || 40));
        await sleep(3000);
        const img = await pv.webContents.capturePage(undefined, { stayHidden: true });
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

let spellLang = 'pt-BR';

const spellHelper = createLineHelper({ timeout: 8000, skipLine: (line) => line.startsWith('init\t'), spawn: () => {
  if (process.platform !== 'win32') return null;
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
  const encoded = Buffer.from(script, 'utf16le').toString('base64');
  return require('child_process').spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded], { windowsHide: true, stdio: ['pipe', 'pipe', 'ignore'] });
} });

function startSpellHelper() { spellHelper.start(); }

function spellQuery(line) {
  return spellHelper.query(line);
}

trustedIpc.handle('spell:suggest', async (_e, word) => {
  const out = await spellQuery(String(word || ''));
  const parts = out.split('\t');
  if (parts[0] !== '1' && parts[0] !== '0') return { misspelled: false, suggestions: [] };
  return { misspelled: parts[0] === '1', suggestions: parts.slice(1).filter(Boolean) };
});

trustedIpc.handle('spell:add-word', async (_e, word) => {
  await spellQuery('+' + String(word || ''));
  try { session.defaultSession.addWordToSpellCheckerDictionary(word); } catch { /* ignora */ }
  return true;
});

app.on('will-quit', () => {
  spellHelper.stop();
  ocrHelper.stop();
});

// ---------------------------------------------------------------- reconhecimento de texto
// PDF digitalizado: o texto das páginas é reconhecido pelo OCR do próprio
// Windows, o mesmo do app Fotos, num processo auxiliar que fica aberto.

const OCR_PS = [
  '[Console]::InputEncoding = [Text.Encoding]::UTF8',
  '[Console]::OutputEncoding = [Text.Encoding]::UTF8',
  "$ErrorActionPreference = 'Stop'",
  'Add-Type -AssemblyName System.Runtime.WindowsRuntime',
  '$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]',
  '$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]',
  '$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]',
  '$null = [Windows.Globalization.Language, Windows.Foundation, ContentType = WindowsRuntime]',
  "$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]",
  'function Await($op, [Type]$t) { $task = $asTask.MakeGenericMethod($t).Invoke($null, @($op)); $task.Wait(-1) | Out-Null; $task.Result }',
  "function J([string]$s) { '\"' + ($s -replace '\\\\', '\\\\' -replace '\"', '\\\"' -replace \"[\\r\\n\\t]\", ' ') + '\"' }",
  "$engine = $null",
  "try { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage([Windows.Globalization.Language]::new('pt-BR')) } catch { }",
  'if (-not $engine) { try { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() } catch { } }',
  'while ($true) {',
  '  $p = [Console]::In.ReadLine(); if ($p -eq $null) { break }',
  '  try {',
  "    if (-not $engine) { throw 'sem-motor' }",
  '    $file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($p)) ([Windows.Storage.StorageFile])',
  '    $stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])',
  '    $decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])',
  '    $bmp = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])',
  '    $res = Await ($engine.RecognizeAsync($bmp)) ([Windows.Media.Ocr.OcrResult])',
  '    $sb = New-Object System.Text.StringBuilder',
  "    [void]$sb.Append('{\"w\":' + $bmp.PixelWidth + ',\"h\":' + $bmp.PixelHeight + ',\"lines\":[')",
  '    $fl = $true',
  '    foreach ($line in $res.Lines) {',
  "      if (-not $fl) { [void]$sb.Append(',') }; $fl = $false; [void]$sb.Append('['); $fw = $true",
  '      foreach ($w in $line.Words) {',
  "        if (-not $fw) { [void]$sb.Append(',') }; $fw = $false; $r = $w.BoundingRect",
  "        [void]$sb.Append('[' + (J $w.Text) + ',' + [math]::Round($r.X, 1) + ',' + [math]::Round($r.Y, 1) + ',' + [math]::Round($r.Width, 1) + ',' + [math]::Round($r.Height, 1) + ']')",
  '      }',
  "      [void]$sb.Append(']')",
  '    }',
  "    [void]$sb.Append(']}')",
  '    $stream.Dispose()',
  '    $o = $sb.ToString()',
  "  } catch { $o = '{\"err\":' + (J $_.Exception.Message) + '}' }",
  '  [Console]::Out.WriteLine($o); [Console]::Out.Flush()',
  '}',
].join('\n');

const ocrHelper = createLineHelper({ timeout: 90000, spawn: () => {
  if (process.platform !== 'win32') return null;
  const encoded = Buffer.from(OCR_PS, 'utf16le').toString('base64');
  return require('child_process').spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', encoded], { windowsHide: true, stdio: ['pipe', 'pipe', 'ignore'] });
} });

function ocrQuery(file) {
  return ocrHelper.query(file);
}

trustedIpc.handle('ocr:image', async (_e, data, ext) => {
  if (process.platform !== 'win32') return { err: 'O reconhecimento de texto só existe no Windows.' };
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'redator-ocr-'));
  const file = path.join(dir, 'pagina.' + (ext === 'png' ? 'png' : 'jpg'));
  try {
    fs.writeFileSync(file, Buffer.from(data));
    const out = await ocrQuery(file);
    if (!out) return { err: 'O reconhecimento de texto não respondeu.' };
    const res = JSON.parse(out);
    if (res.err === 'sem-motor') res.err = 'O Windows não tem o reconhecimento de texto em português instalado.';
    return res;
  } catch (err) {
    return { err: err.message };
  } finally {
    try { fs.rmSync(dir, { recursive: true, force: true }); } catch { /* ignora */ }
  }
});

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
