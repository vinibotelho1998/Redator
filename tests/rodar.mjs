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
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'redator-teste-'));
const fixtures = path.join(root, 'tests', 'fixtures') + path.sep;
const script = 'window.__fixtures = ' + JSON.stringify(fixtures) + ';\nwindow.__ci = ' + JSON.stringify(!!process.env.CI) + ';\n' + fs.readFileSync(path.join(root, 'tests', 'suite.js'), 'utf8');
const scriptPath = path.join(tmp, 'suite.js');
fs.writeFileSync(scriptPath, script);

const run = (cmd, args, env) => new Promise((resolve) => {
  const child = spawn(cmd, args, { cwd: root, env: { ...process.env, ...env }, stdio: ['ignore', 'pipe', 'pipe'] });
  let buf = '';
  let done = null;
  const timer = setTimeout(() => { console.error('Tempo esgotado.'); child.kill(); }, 5 * 60 * 1000);
  const onData = (chunk) => {
    buf += chunk.toString('utf8');
    let i;
    while ((i = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, i).replace(/\r$/, '');
      buf = buf.slice(i + 1);
      if (line.startsWith('LOG ')) {
        console.log(line.slice(4));
        const m = /^LOG FIM ok=(\d+) falhas=(\d+)/.exec(line);
        if (m) done = { ok: Number(m[1]), fail: Number(m[2]) };
      } else if (/CONSOLE-ERR|TEST-ERR/.test(line)) console.log(line);
    }
  };
  child.stdout.on('data', onData);
  child.stderr.on('data', () => {});
  child.on('exit', () => { clearTimeout(timer); resolve(done); });
});

await new Promise((resolve, reject) => {
  const b = spawn(process.execPath, ['build.mjs'], { cwd: root, stdio: 'inherit' });
  b.on('exit', (code) => (code === 0 ? resolve() : reject(new Error('build falhou'))));
});

const result = await run(electronPath, [root], {
  REDATOR_TEST: scriptPath,
  REDATOR_TEST_OUT: tmp,
  REDATOR_TEST_DELAY: '1000',
  ELECTRON_ENABLE_LOGGING: '',
});
try { fs.rmSync(tmp, { recursive: true, force: true }); } catch { /* ignora */ }
if (!result) {
  console.error('A bateria não terminou.');
  process.exit(1);
}
console.log('\n' + result.ok + ' testes passaram, ' + result.fail + ' falharam.');
process.exit(result.fail ? 1 : 0);
