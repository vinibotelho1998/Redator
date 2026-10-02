// Roda a bateria de testes: sobe o Redator em modo de teste, executa o
// tests/suite.js dentro da janela e devolve código de saída 1 se algo falhar.
//
//   npm test
import { spawn } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import electronPath from 'electron';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const packaged = process.argv.includes('--packaged');
const executable = packaged ? path.join(root, 'out', 'Redator-win32-x64', 'Redator.exe') : electronPath;
if (!fs.existsSync(executable)) throw new Error('Executável não encontrado: ' + executable);
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'redator-teste-'));
const fixtures = path.join(root, 'tests', 'fixtures') + path.sep;
const resultsDir = path.join(root, 'tests-results');
fs.mkdirSync(resultsDir, { recursive: true });
const logPath = path.join(resultsDir, packaged ? 'electron-packaged.log' : 'electron.log');
fs.writeFileSync(logPath, '');
const regressions = ['ui-regressions.js', 'editor-regressions.js', 'document-regressions.js']
  .map((name) => { const p = path.join(root, 'tests', name); return fs.existsSync(p) ? fs.readFileSync(p, 'utf8') : ''; }).join('\n');
const script = 'window.__fixtures = ' + JSON.stringify(fixtures) + ';\nwindow.__tmp = ' + JSON.stringify(tmp + path.sep) + ';\nwindow.__ci = ' + JSON.stringify(!!process.env.CI) + ';\n' + regressions + '\n' + fs.readFileSync(path.join(root, 'tests', 'suite.js'), 'utf8');
const scriptPath = path.join(tmp, 'suite.js');
fs.writeFileSync(scriptPath, script);

const run = (cmd, args, env) => new Promise((resolve) => {
  const childEnv = { ...process.env, ...env };
  delete childEnv.ELECTRON_RUN_AS_NODE;
  const child = spawn(cmd, args, { cwd: root, env: childEnv, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
  let buf = '';
  let done = null;
  let stderr = '';
  const timer = setTimeout(() => { console.error('Tempo esgotado.'); child.kill(); }, 5 * 60 * 1000);
  const onData = (chunk) => {
    buf += chunk.toString('utf8');
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i).replace(/\r$/, '');
      buf = buf.slice(i + 1);
      fs.appendFileSync(logPath, line + '\n');
      if (line.startsWith('LOG ')) {
        console.log(line.slice(4));
        const m = /^LOG FIM ok=(\d+) falhas=(\d+)/.exec(line);
        if (m) done = { ok: Number(m[1]), fail: Number(m[2]) };
      } else if (/CONSOLE-ERR|TEST-ERR/.test(line)) console.log(line);
    }
  };
  child.stdout.on('data', onData);
  child.stderr.on('data', (chunk) => {
    const message = chunk.toString('utf8');
    stderr = (stderr + message).slice(-12000);
    fs.appendFileSync(logPath, message);
  });
  child.on('error', (err) => {
    clearTimeout(timer);
    console.error('Não foi possível iniciar o Electron:', err.message);
    resolve(null);
  });
  child.on('exit', (code, signal) => {
    clearTimeout(timer);
    if (!done || code !== 0 || signal) {
      console.error('Electron terminou:', { code, signal });
      if (stderr.trim()) console.error(stderr.trim());
      resolve(null);
    } else resolve(done);
  });
});

let result;
try {
  if (!packaged) await new Promise((resolve, reject) => {
    const b = spawn(process.execPath, ['build.mjs'], { cwd: root, stdio: 'inherit', windowsHide: true });
    b.on('error', reject);
    b.on('exit', (code) => (code === 0 ? resolve() : reject(new Error('build falhou'))));
  });
  result = await run(executable, packaged ? [] : [root], {
    REDATOR_TEST: scriptPath,
    REDATOR_TEST_OUT: tmp,
    REDATOR_TEST_DELAY: '1000',
    ELECTRON_ENABLE_LOGGING: '',
  });
} finally {
  try { fs.rmSync(tmp, { recursive: true, force: true }); } catch { /* ignora */ }
}
if (!result) {
  console.error('A bateria não terminou.');
  process.exit(1);
}
console.log('\n' + result.ok + ' testes passaram, ' + result.fail + ' falharam.');
process.exit(result.fail ? 1 : 0);
