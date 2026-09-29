'use strict';

const { contextBridge, ipcRenderer, webUtils } = require('electron');

const CHANNELS = ['app:request-close', 'app:context-menu', 'app:window-state', 'app:focus', 'app:open-file'];

contextBridge.exposeInMainWorld('host', {
  isElectron: true,
  platform: process.platform,
  init: () => ipcRenderer.invoke('app:init'),
  readFile: (p) => ipcRenderer.invoke('file:read', p),
  fileExists: (p) => ipcRenderer.invoke('file:exists', p),
  writeFile: (p, data, extras) => ipcRenderer.invoke('file:write', p, data, extras),
  openDialog: (opts) => ipcRenderer.invoke('file:open-dialog', opts),
  imageDialog: () => ipcRenderer.invoke('file:image-dialog'),
  saveDialog: (opts) => ipcRenderer.invoke('file:save-dialog', opts),
  setDocState: (st) => ipcRenderer.send('app:set-doc-state', st),
  recent: () => ipcRenderer.invoke('app:recent'),
  addRecent: (p) => ipcRenderer.invoke('app:recent-add', p),
  removeRecent: (p) => ipcRenderer.invoke('app:recent-remove', p),
  newWindow: (p) => ipcRenderer.invoke('app:new-window', p || null),
  closeConfirmed: () => ipcRenderer.send('app:close-confirmed'),
  closeWindow: () => ipcRenderer.send('app:close-window'),
  quit: () => ipcRenderer.send('app:quit'),
  setOverlay: (colors) => ipcRenderer.send('app:set-overlay', colors),
  savePrefs: (prefs) => ipcRenderer.invoke('app:save-prefs', prefs),
  print: () => ipcRenderer.invoke('app:print'),
  exportPdf: (p, opts) => ipcRenderer.invoke('app:export-pdf', p, opts),
  openExternal: (url) => ipcRenderer.invoke('shell:open-external', url),
  openPath: (p) => ipcRenderer.invoke('shell:open-path', p),
  showInFolder: (p) => ipcRenderer.invoke('shell:show-item', p),
  replaceMisspelling: (w) => ipcRenderer.send('spell:replace', w),
  addToDictionary: (w) => ipcRenderer.send('spell:add', w),
  spellLanguages: () => ipcRenderer.invoke('spell:languages'),
  setSpellLanguages: (langs) => ipcRenderer.invoke('spell:set-languages', langs),
  systemFonts: () => ipcRenderer.invoke('fonts:list'),
  convertImage: (data, ext, w, h) => ipcRenderer.invoke('image:convert', data, ext, w, h),
  pathForFile: (file) => { try { return webUtils.getPathForFile(file); } catch { return ''; } },
  edit: (action) => ipcRenderer.send('edit:action', action),
  findNote: (from, target) => ipcRenderer.invoke('file:find-note', from, target),
  ocr: (data, ext) => ipcRenderer.invoke('ocr:image', data, ext),
  defaultStatus: () => ipcRenderer.invoke('app:default-status'),
  openDefaultApps: () => ipcRenderer.invoke('app:open-default-apps'),
  saveRecovery: (data, meta) => ipcRenderer.invoke('recover:save', data, meta),
  clearRecovery: () => ipcRenderer.send('recover:clear'),
  takeRecovery: (id) => ipcRenderer.invoke('recover:take', id),
  openRecovered: (id) => ipcRenderer.invoke('recover:open', id),
  discardRecovery: (ids) => ipcRenderer.invoke('recover:discard', ids),
  checkWord: (word) => ipcRenderer.invoke('spell:suggest', word),
  addWord: (word) => ipcRenderer.invoke('spell:add-word', word),
  on: (channel, fn) => {
    if (!CHANNELS.includes(channel)) return;
    ipcRenderer.on(channel, (_e, ...args) => fn(...args));
  },
});
