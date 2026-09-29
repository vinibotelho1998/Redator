// Ponte com o processo principal. No navegador, usa alternativas simples para testes.

class BrowserHost {
  constructor() {
    this.isElectron = false;
    this.platform = 'browser';
    this.files = new Map();
  }
  async init() {
    const q = new URLSearchParams(location.search);
    const prefs = {};
    if (q.get('zoom')) prefs.zoom = Number(q.get('zoom'));
    if (q.get('marks')) prefs.showMarks = true;
    if (q.get('page')) prefs.testPage = Number(q.get('page'));
    if (q.get('nav')) prefs.navPane = true;
    return { filePath: q.get('file'), recent: [], version: 'dev', userName: '', prefs, platform: 'browser' };
  }
  async readFile(p) {
    if (this.files.has(p)) {
      const f = this.files.get(p);
      return { data: new Uint8Array(await f.arrayBuffer()), readonly: false, name: f.name };
    }
    const data = await new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      xhr.open('GET', p);
      xhr.responseType = 'arraybuffer';
      xhr.onload = () => (xhr.status === 0 || xhr.status === 200 ? resolve(new Uint8Array(xhr.response)) : reject(new Error('Status ' + xhr.status)));
      xhr.onerror = () => reject(new Error('Não foi possível ler ' + p));
      xhr.send();
    });
    return { data, readonly: false, name: decodeURIComponent(p.split(/[\\/]/).pop()) };
  }
  async fileExists() { return false; }
  pick(accept, multiple) {
    return new Promise((resolve) => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = accept || '';
      input.multiple = !!multiple;
      input.onchange = () => resolve(Array.from(input.files || []));
      input.click();
    });
  }
  async openDialog() {
    const files = await this.pick('.docx,.odt,.md,.markdown,.txt,.html,.htm', true);
    return files.map((f) => { const key = 'local:' + f.name; this.files.set(key, f); return key; });
  }
  async imageDialog() {
    const files = await this.pick('image/*', true);
    return Promise.all(files.map(async (f) => ({ path: f.name, name: f.name, data: new Uint8Array(await f.arrayBuffer()) })));
  }
  async saveDialog(opts) { return (opts && opts.defaultPath) || 'documento'; }
  async writeFile(p, data) {
    const blob = new Blob([data]);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = String(p).split(/[\\/]/).pop();
    a.click();
    return { ok: true };
  }
  setDocState(st) { if (st.title) document.title = st.title; }
  async recent() { return []; }
  async addRecent() { return []; }
  async removeRecent() { return []; }
  async newWindow(p) { window.open(location.pathname + (p ? '?file=' + encodeURIComponent(p) : ''), '_blank'); }
  closeConfirmed() { window.close(); }
  closeWindow() { window.close(); }
  quit() {}
  setOverlay() {}
  async savePrefs(p) { return p; }
  async print() { window.print(); return { ok: true }; }
  async exportPdf() { window.print(); return { ok: true }; }
  async openExternal(url) { window.open(url, '_blank'); }
  async openPath() {}
  async showInFolder() {}
  replaceMisspelling() {}
  addToDictionary() {}
  async spellLanguages() { return { available: ['pt-BR', 'en-US'], current: ['pt-BR'] }; }
  async setSpellLanguages() { return true; }
  async systemFonts() { return []; }
  async convertImage() { return null; }
  pathForFile(f) { const key = 'local:' + f.name; this.files.set(key, f); return key; }
  edit(action) { document.execCommand(action === 'pasteAndMatchStyle' ? 'paste' : action); }
  on() {}
}

export function getHost() {
  if (window.host && window.host.isElectron) {
    const h = window.host;
    return h;
  }
  return new BrowserHost();
}
