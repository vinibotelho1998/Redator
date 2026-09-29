// Cores: realce, cores de tema, tonalidade e sombreamento.

export const HIGHLIGHT = {
  yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0', black: '#000000', white: '#FFFFFF',
};

export const HIGHLIGHT_LIST = [
  ['yellow', 'Amarelo'], ['green', 'Verde Brilhante'], ['cyan', 'Turquesa'], ['magenta', 'Rosa'], ['blue', 'Azul'],
  ['red', 'Vermelho'], ['darkBlue', 'Azul-escuro'], ['darkCyan', 'Azul-petróleo'], ['darkGreen', 'Verde'],
  ['darkMagenta', 'Violeta'], ['darkRed', 'Vermelho-escuro'], ['darkYellow', 'Amarelo-escuro'],
  ['darkGray', 'Cinza 50%'], ['lightGray', 'Cinza 25%'], ['black', 'Preto'],
];

export const DEFAULT_THEME_COLORS = {
  dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
  accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
  hlink: '0563C1', folHlink: '954F72',
};

const THEME_NAME_MAP = {
  dark1: 'dk1', light1: 'lt1', dark2: 'dk2', light2: 'lt2', text1: 'dk1', background1: 'lt1', text2: 'dk2', background2: 'lt2',
  tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2',
  accent1: 'accent1', accent2: 'accent2', accent3: 'accent3', accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
  hyperlink: 'hlink', followedHyperlink: 'folHlink', hlink: 'hlink', folHlink: 'folHlink',
  dk1: 'dk1', lt1: 'lt1', dk2: 'dk2', lt2: 'lt2',
};

export function themeKey(name) {
  return THEME_NAME_MAP[name] || name;
}

export function normHex(v) {
  if (!v) return null;
  const s = String(v).trim().replace(/^#/, '');
  if (/^[0-9a-f]{6}$/i.test(s)) return '#' + s.toUpperCase();
  if (/^[0-9a-f]{3}$/i.test(s)) return '#' + s.split('').map((c) => c + c).join('').toUpperCase();
  return null;
}

export function hexToRgb(h) {
  const s = normHex(h);
  if (!s) return [0, 0, 0];
  return [parseInt(s.slice(1, 3), 16), parseInt(s.slice(3, 5), 16), parseInt(s.slice(5, 7), 16)];
}

export function rgbToHex([r, g, b]) {
  const c = (x) => Math.max(0, Math.min(255, Math.round(x))).toString(16).padStart(2, '0');
  return ('#' + c(r) + c(g) + c(b)).toUpperCase();
}

export function rgbToHsl([r, g, b]) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  return [h, s, l];
}

export function hslToRgb([h, s, l]) {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const hue = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [hue(p, q, h + 1 / 3) * 255, hue(p, q, h) * 255, hue(p, q, h - 1 / 3) * 255];
}

// w:themeTint e w:themeShade em hexadecimal 00 a FF
export function applyTintShade(hex, tint, shade) {
  let hsl = rgbToHsl(hexToRgb(hex));
  if (tint != null) {
    const t = parseInt(tint, 16) / 255;
    hsl = [hsl[0], hsl[1], hsl[2] * t + (1 - t)];
  }
  if (shade != null) {
    const s = parseInt(shade, 16) / 255;
    hsl = [hsl[0], hsl[1], hsl[2] * s];
  }
  return rgbToHex(hslToRgb(hsl));
}

// Modificadores DrawingML: lumMod, lumOff em milésimos de por cento
export function applyLum(hex, lumMod, lumOff) {
  let [h, s, l] = rgbToHsl(hexToRgb(hex));
  if (lumMod != null) l = l * (lumMod / 100000);
  if (lumOff != null) l = l + lumOff / 100000;
  return rgbToHex(hslToRgb([h, s, Math.max(0, Math.min(1, l))]));
}

export function mix(hexA, hexB, t) {
  const a = hexToRgb(hexA), b = hexToRgb(hexB);
  return rgbToHex([a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t, a[2] + (b[2] - a[2]) * t]);
}

export function isDark(hex) {
  const [r, g, b] = hexToRgb(hex);
  return (0.299 * r + 0.587 * g + 0.114 * b) < 110;
}

// Paleta de cores de tema como no Word: 10 colunas, cor base + 5 variações
export function themePalette(colors) {
  const base = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
  const names = ['Branco, Plano de Fundo 1', 'Preto, Texto 1', 'Cinza-claro, Plano de Fundo 2', 'Azul-escuro, Texto 2',
    'Ênfase 1', 'Ênfase 2', 'Ênfase 3', 'Ênfase 4', 'Ênfase 5', 'Ênfase 6'];
  const themeNames = ['background1', 'text1', 'background2', 'text2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
  return base.map((k, i) => {
    const hex = normHex(colors[k]) || '#000000';
    const [, , l] = rgbToHsl(hexToRgb(hex));
    let variants;
    if (l < 0.2) variants = [['80', null, 'Mais Claro 50%'], ['A6', null, 'Mais Claro 35%'], ['BF', null, 'Mais Claro 25%'], ['D9', null, 'Mais Claro 15%'], ['F2', null, 'Mais Claro 5%']];
    else if (l > 0.8) variants = [[null, 'F2', 'Mais Escuro 5%'], [null, 'D9', 'Mais Escuro 15%'], [null, 'BF', 'Mais Escuro 25%'], [null, 'A6', 'Mais Escuro 35%'], [null, '80', 'Mais Escuro 50%']];
    else variants = [['33', null, 'Mais Claro 80%'], ['66', null, 'Mais Claro 60%'], ['99', null, 'Mais Claro 40%'], [null, 'BF', 'Mais Escuro 25%'], [null, '80', 'Mais Escuro 50%']];
    return {
      name: names[i], themeColor: themeNames[i], hex,
      variants: variants.map(([tint, shade, label]) => ({
        hex: applyTintShade(hex, tint, shade),
        tint, shade, label: names[i] + ', ' + label, themeColor: themeNames[i],
      })),
    };
  });
}

export const STANDARD_COLORS = [
  ['#C00000', 'Vermelho-escuro'], ['#FF0000', 'Vermelho'], ['#FFC000', 'Laranja'], ['#FFFF00', 'Amarelo'], ['#92D050', 'Verde-claro'],
  ['#00B050', 'Verde'], ['#00B0F0', 'Azul-claro'], ['#0070C0', 'Azul'], ['#002060', 'Azul-escuro'], ['#7030A0', 'Roxo'],
];
