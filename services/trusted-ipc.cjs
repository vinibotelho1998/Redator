'use strict';

function createTrustedIpc(ipcMain, { windows, getWindow, appUrl }) {
  function allowed(event) {
    try {
      const win = getWindow(event);
      return !!win && windows.has(win) && !win.isDestroyed()
        && event.senderFrame === event.sender.mainFrame
        && event.senderFrame.url === appUrl;
    } catch { return false; }
  }
  return {
    handle(channel, listener) {
      ipcMain.handle(channel, (event, ...args) => {
        if (!allowed(event)) throw new Error('Origem IPC não autorizada.');
        return listener(event, ...args);
      });
    },
    on(channel, listener) {
      ipcMain.on(channel, (event, ...args) => {
        if (!allowed(event)) return;
        try {
          const result = listener(event, ...args);
          if (result && typeof result.catch === 'function') result.catch((err) => console.warn('Falha na operação ' + channel, err.message));
        }
        catch (err) { console.warn('Falha na operação ' + channel, err.message); }
      });
    },
  };
}

module.exports = { createTrustedIpc };
