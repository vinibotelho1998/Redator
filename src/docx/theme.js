// Tema do documento: cores e fontes.
import { parseXml, NS } from '../util/xml.js';
import { DEFAULT_THEME_COLORS, themeKey, normHex, applyTintShade } from '../util/color.js';

export class Theme {
  constructor() {
    this.colors = { ...DEFAULT_THEME_COLORS };
    this.major = 'Calibri Light';
    this.minor = 'Calibri';
    this.majorEa = '';
    this.minorEa = '';
    this.majorCs = '';
    this.minorCs = '';
    this.xml = null;
  }

  hex(name, tint, shade) {
    const key = themeKey(name);
    const base = normHex(this.colors[key]);
    if (!base) return null;
    if (tint || shade) return applyTintShade(base, tint || null, shade || null);
    return base;
  }

  font(ref) {
    if (!ref) return null;
    const major = ref.startsWith('major');
    if (/EastAsia$/.test(ref)) return (major ? this.majorEa : this.minorEa) || (major ? this.major : this.minor);
    if (/Bidi$/.test(ref)) return (major ? this.majorCs : this.minorCs) || (major ? this.major : this.minor);
    return major ? this.major : this.minor;
  }
}

export function parseTheme(text) {
  const t = new Theme();
  if (!text) return t;
  t.xml = text;
  let doc;
  try { doc = parseXml(text); } catch { return t; }
  const clr = doc.getElementsByTagNameNS(NS.a, 'clrScheme')[0];
  if (clr) {
    for (const c of clr.children) {
      const srgb = c.getElementsByTagNameNS(NS.a, 'srgbClr')[0];
      const sys = c.getElementsByTagNameNS(NS.a, 'sysClr')[0];
      let v = null;
      if (srgb) v = srgb.getAttribute('val');
      else if (sys) v = sys.getAttribute('lastClr') || (sys.getAttribute('val') === 'window' ? 'FFFFFF' : '000000');
      if (v) t.colors[c.localName] = v;
    }
  }
  const fs = doc.getElementsByTagNameNS(NS.a, 'fontScheme')[0];
  if (fs) {
    const read = (el, tag) => {
      const x = el && el.getElementsByTagNameNS(NS.a, tag)[0];
      return x ? x.getAttribute('typeface') || '' : '';
    };
    const maj = fs.getElementsByTagNameNS(NS.a, 'majorFont')[0];
    const min = fs.getElementsByTagNameNS(NS.a, 'minorFont')[0];
    t.major = read(maj, 'latin') || t.major;
    t.minor = read(min, 'latin') || t.minor;
    t.majorEa = read(maj, 'ea');
    t.minorEa = read(min, 'ea');
    t.majorCs = read(maj, 'cs');
    t.minorCs = read(min, 'cs');
  }
  return t;
}
