// Gera o ícone do Redator (ICO e PNG) renderizando um SVG com o próprio Electron.
// Uso: electron tools/make-icon.cjs
const { app, BrowserWindow, nativeImage } = require('electron');
const fs = require('fs');
const path = require('path');

const SVG = `<svg xmlns="http://www.w3.org/2000/svg" width="256" height="256" viewBox="0 0 256 256">
  <defs>
    <linearGradient id="g" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#2f7fd8"/><stop offset="1" stop-color="#0f4495"/></linearGradient>
    <linearGradient id="p" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#ffffff"/><stop offset="1" stop-color="#e9f0fb"/></linearGradient>
  </defs>
  <rect x="78" y="22" width="156" height="212" rx="16" fill="url(#p)" stroke="#b9cdee" stroke-width="4"/>
  <path d="M112 70h94M112 102h94M112 134h94M112 166h94M112 198h60" stroke="#8fb2e6" stroke-width="11" stroke-linecap="round"/>
  <rect x="18" y="58" width="142" height="142" rx="22" fill="url(#g)"/>
  <text x="89" y="167" font-family="Segoe UI, Arial" font-weight="700" font-size="116" fill="#ffffff" text-anchor="middle">R</text>
</svg>`;

function ico(pngs) {
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);
  header.writeUInt16LE(1, 2);
  header.writeUInt16LE(pngs.length, 4);
  const dir = Buffer.alloc(16 * pngs.length);
  let offset = 6 + dir.length;
  pngs.forEach(({ size, data }, i) => {
    const o = i * 16;
    dir.writeUInt8(size >= 256 ? 0 : size, o);
    dir.writeUInt8(size >= 256 ? 0 : size, o + 1);
    dir.writeUInt8(0, o + 2);
    dir.writeUInt8(0, o + 3);
    dir.writeUInt16LE(1, o + 4);
    dir.writeUInt16LE(32, o + 6);
    dir.writeUInt32LE(data.length, o + 8);
    dir.writeUInt32LE(offset, o + 12);
    offset += data.length;
  });
  return Buffer.concat([header, dir, ...pngs.map((p) => p.data)]);
}

app.whenReady().then(async () => {
  const win = new BrowserWindow({ width: 256, height: 256, show: false, frame: false, transparent: true, webPreferences: { offscreen: true } });
  const html = '<html><body style="margin:0;background:transparent;overflow:hidden">' + SVG + '</body></html>';
  await win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(html));
  await new Promise((r) => setTimeout(r, 800));
  const img = await win.webContents.capturePage({ x: 0, y: 0, width: 256, height: 256 });
  const outDir = path.join(__dirname, '..', 'app', 'icons');
  fs.mkdirSync(outDir, { recursive: true });
  const sizes = [256, 128, 64, 48, 32, 24, 16];
  const pngs = sizes.map((size) => ({ size, data: (size === 256 ? img : img.resize({ width: size, height: size, quality: 'best' })).toPNG() }));
  fs.writeFileSync(path.join(outDir, 'redator.ico'), ico(pngs));
  fs.writeFileSync(path.join(outDir, 'redator.png'), pngs[0].data);
  fs.writeFileSync(path.join(outDir, 'redator-64.png'), pngs[2].data);
  console.log('ícone gerado em', outDir);
  void nativeImage;
  app.quit();
});
