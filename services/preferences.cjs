'use strict';

function normalizePreferences(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return {};
  const next = {};
  for (const key of ['spellcheck', 'smartQuotes', 'autoFormat', 'autoRecover', 'ruler', 'navPane', 'showMarks']) {
    if (typeof value[key] === 'boolean') next[key] = value[key];
  }
  if (['intl', 'pt'].includes(value.shortcuts)) next.shortcuts = value.shortcuts;
  if (['cm', 'in', 'pt'].includes(value.units)) next.units = value.units;
  if (typeof value.author === 'string') next.author = value.author.slice(0, 1000);
  if (Number.isFinite(value.zoom) && value.zoom >= 0.1 && value.zoom <= 5) next.zoom = value.zoom;
  if (Number.isInteger(value.testPage) && value.testPage > 0) next.testPage = value.testPage;
  for (const [key, max, length] of [['recentFonts', 8, 200], ['recentSymbols', 20, 32], ['spellLanguages', 20, 40]]) {
    if (Array.isArray(value[key])) next[key] = [...new Set(value[key].filter((item) => typeof item === 'string' && item.length && item.length <= length))].slice(0, max);
  }
  return next;
}

module.exports = { normalizePreferences };
