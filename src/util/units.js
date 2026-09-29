// Conversões de unidades. Base de tela: 96 px por polegada.

export const twipToPx = (t) => t / 15;
export const pxToTwip = (p) => Math.round(p * 15);
export const twipToPt = (t) => t / 20;
export const ptToTwip = (p) => Math.round(p * 20);
export const ptToPx = (p) => (p * 4) / 3;
export const pxToPt = (p) => (p * 3) / 4;
export const emuToPx = (e) => e / 9525;
export const pxToEmu = (p) => Math.round(p * 9525);
export const emuToPt = (e) => e / 12700;
export const ptToEmu = (p) => Math.round(p * 12700);
export const TWIP_PER_CM = 1440 / 2.54;
export const twipToCm = (t) => t / TWIP_PER_CM;
export const cmToTwip = (c) => Math.round(c * TWIP_PER_CM);

export function num(v, def = 0) {
  if (v == null || v === '') return def;
  const n = typeof v === 'number' ? v : parseFloat(v);
  return Number.isFinite(n) ? n : def;
}

export function fmtNumber(n, digits = 2) {
  const f = Number(n.toFixed(digits));
  return String(f).replace('.', ',');
}

export function fmtCm(twips, digits = 2) {
  return fmtNumber(twipToCm(twips), digits) + ' cm';
}

export function fmtPt(pt, digits = 1) {
  return fmtNumber(pt, digits) + ' pt';
}

// Aceita "2,5", "2.5 cm", "1 pol", "12 pt", "1\"" e devolve twips
export function parseLength(str, defUnit = 'cm') {
  if (str == null) return null;
  const s = String(str).trim().toLowerCase().replace(',', '.');
  const m = /^(-?\d*\.?\d+)\s*(cm|mm|pt|pol|in|"|px|li|linha|linhas)?$/.exec(s);
  if (!m) return null;
  const v = parseFloat(m[1]);
  const u = m[2] || defUnit;
  switch (u) {
    case 'cm': return Math.round(v * TWIP_PER_CM);
    case 'mm': return Math.round((v / 10) * TWIP_PER_CM);
    case 'pt': return Math.round(v * 20);
    case 'pol': case 'in': case '"': return Math.round(v * 1440);
    case 'px': return Math.round(v * 15);
    case 'li': case 'linha': case 'linhas': return Math.round(v * 240);
    default: return Math.round(v * TWIP_PER_CM);
  }
}

export function clamp(v, lo, hi) {
  return Math.max(lo, Math.min(hi, v));
}
