'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { normalizePreferences } = require('../services/preferences.cjs');

test('preferências válidas conservam unidades, zoom, idiomas e falsos explícitos', () => {
  const valid = { units: 'in', shortcuts: 'pt', zoom: 1.25, author: 'Vinícius', autoRecover: false, spellcheck: true, recentFonts: ['Arial'], recentSymbols: ['✦'], spellLanguages: ['pt-BR', 'en-US'] };
  assert.deepEqual(normalizePreferences(valid), valid);
});

test('preferências corrompidas não substituem defaults por tipos inválidos', () => {
  assert.deepEqual(normalizePreferences(null), {});
  assert.deepEqual(normalizePreferences([]), {});
  assert.deepEqual(normalizePreferences({ units: {}, shortcuts: 'errado', zoom: '1', ruler: 'false', navPane: {}, recentFonts: 'Arial', author: [], spellLanguages: false, __proto__: { dangerous: true } }), {});
  assert.deepEqual(normalizePreferences({ zoom: Infinity, recentSymbols: [null, {}, '', '✓', '✓'], recentFonts: ['Arial', 10, 'Arial', 'Calibri'] }), { recentFonts: ['Arial', 'Calibri'], recentSymbols: ['✓'] });
  assert.deepEqual(normalizePreferences({ zoom: 99 }), {});
});
