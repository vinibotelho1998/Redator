'use strict';

// The helpers speak one response per input line. Once a response times out,
// restart the process so a late response cannot be assigned to another query.
function createLineHelper({ spawn, timeout, skipLine = () => false, maxBuffer = 4 * 1024 * 1024 }) {
  let current = null;

  function stop(state = current) {
    if (!state || state.dead) return;
    state.dead = true;
    if (current === state) current = null;
    for (const job of state.jobs.splice(0)) job.finish('');
    try { state.proc.kill(); } catch { /* already exited */ }
  }

  function start() {
    if (current) return current;
    let proc;
    try { proc = spawn(); } catch { return null; }
    if (!proc) return null;
    const state = { proc, buffer: '', jobs: [], dead: false };
    current = state;
    proc.on('error', () => stop(state));
    proc.on('exit', () => stop(state));
    proc.on('close', () => stop(state));
    proc.stdin.on('error', () => stop(state));
    proc.stdout.on('error', () => stop(state));
    proc.stdout.setEncoding('utf8');
    proc.stdout.on('data', (chunk) => {
      if (state.dead) return;
      state.buffer += chunk;
      if (state.buffer.length > maxBuffer) { stop(state); return; }
      let newline;
      while ((newline = state.buffer.indexOf('\n')) >= 0) {
        const line = state.buffer.slice(0, newline).replace(/\r$/, '');
        state.buffer = state.buffer.slice(newline + 1);
        if (skipLine(line)) continue;
        const job = state.jobs.shift();
        if (job) job.finish(line);
      }
    });
    return state;
  }

  function query(line) {
    const state = start();
    if (!state) return Promise.resolve('');
    return new Promise((resolve) => {
      const job = { finish: (out) => { clearTimeout(timer); resolve(out); } };
      const timer = setTimeout(() => stop(state), timeout);
      state.jobs.push(job);
      try {
        state.proc.stdin.write(String(line).replace(/[\r\n]/g, ' ') + '\n', (err) => { if (err) stop(state); });
      } catch { stop(state); }
    });
  }

  return { start, query, stop: () => stop() };
}

module.exports = { createLineHelper };
