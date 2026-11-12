import { spawn } from 'node:child_process';
import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { ViewerError } from './errors.mjs';
import { powershellExecutable } from './powershell.mjs';

const SCRIPT = fileURLToPath(new URL('../scripts/render-deck.ps1', import.meta.url));
const DEFAULT_TIMEOUT_MS = 5 * 60 * 1000;
const FORCE_KILL_AFTER_MS = 15 * 1000;
const MAX_STDERR = 4096;

function describeFailure(stderr) {
  const text = stderr.trim();
  if (/0x80040154|REGDB_E_CLASSNOTREG|Class not registered/i.test(text)) {
    return new ViewerError('powerpoint_unavailable', 'error.powerpointUnavailable');
  }
  const detail = text.replace(/^HRESULT=\S+\s*/, '').split(/\r?\n/)[0].slice(0, 300);
  return detail ? new ViewerError('render_failed', 'error.renderFailedDetail', { detail }) : new ViewerError('render_failed', 'error.renderFailed');
}

/**
 * Renders every slide of `pptx` into `outDir` with desktop PowerPoint and returns the parsed slides.json.
 * `onProgress(done, total)` is called as slides finish. Aborting the signal stops the run between slides.
 */
export async function renderWithPowerPoint({ pptx, outDir, width, thumbWidth, signal, onProgress, timeoutMs = DEFAULT_TIMEOUT_MS }) {
  if (process.platform !== 'win32') {
    throw new ViewerError('unsupported_platform', 'error.renderNeedsWindows');
  }
  const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', SCRIPT, '-Pptx', pptx, '-OutDir', outDir, '-Width', String(width), '-ThumbWidth', String(thumbWidth)];
  const child = spawn(powershellExecutable(), args, { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });

  let stderr = '';
  let pending = '';
  let stopReason = null;
  let killTimer = null;

  child.stderr.setEncoding('utf8');
  child.stderr.on('data', (chunk) => {
    if (stderr.length < MAX_STDERR) stderr += chunk;
  });
  child.stdout.setEncoding('utf8');
  child.stdout.on('data', (chunk) => {
    pending += chunk;
    const lines = pending.split(/\r?\n/);
    pending = lines.pop() ?? '';
    for (const line of lines) {
      const match = /^PROGRESS (\d+) (\d+)$/.exec(line.trim());
      if (match) onProgress?.(Number(match[1]), Number(match[2]));
    }
  });

  const stop = (reason) => {
    if (stopReason) return;
    stopReason = reason;
    fs.writeFile(path.join(outDir, 'cancel.flag'), '1').catch(() => {});
    killTimer = setTimeout(() => child.kill(), FORCE_KILL_AFTER_MS);
  };
  const onAbort = () => stop('cancelled');
  if (signal?.aborted) onAbort();
  else signal?.addEventListener('abort', onAbort, { once: true });
  const timeoutTimer = setTimeout(() => stop('timeout'), timeoutMs);

  const exitCode = await new Promise((resolve, reject) => {
    child.once('error', reject);
    child.once('close', (code) => resolve(code));
  }).catch((error) => {
    throw new ViewerError('render_failed', 'error.powershellStartFailed', { detail: error.message });
  }).finally(() => {
    clearTimeout(timeoutTimer);
    clearTimeout(killTimer);
    signal?.removeEventListener('abort', onAbort);
  });

  if (stopReason === 'timeout') throw new ViewerError('render_timeout', 'error.renderTimeout');
  if (stopReason === 'cancelled' || exitCode === 2) throw new ViewerError('cancelled', 'error.cancelled');
  if (exitCode !== 0) throw describeFailure(stderr);

  let parsed;
  try {
    parsed = JSON.parse(await fs.readFile(path.join(outDir, 'slides.json'), 'utf8'));
  } catch (error) {
    throw new ViewerError('render_failed', 'error.renderResultUnreadable', { detail: error.message });
  }
  return parsed;
}
