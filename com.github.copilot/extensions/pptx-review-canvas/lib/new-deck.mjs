import { spawn } from 'node:child_process';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { resolveDeckSource } from './deck-cache.mjs';
import { ViewerError } from './errors.mjs';
import { tEn } from './i18n.mjs';
import { powershellExecutable } from './powershell.mjs';

const SCRIPT = fileURLToPath(new URL('../scripts/new-deck.ps1', import.meta.url));
const TIMEOUT_MS = 3 * 60 * 1000;
const MAX_STDERR = 4096;

const exists = (target) => fs.access(target).then(() => true, () => false);

/**
 * Takes the file name for good by creating it empty: of two callers asking for the same name, one gets `file_exists`, however
 * the timing falls. PowerPoint then saves over the empty file. Returns a function that removes the file again if it is still empty.
 */
export async function reserveNewDeckFile(file) {
  let handle;
  try {
    handle = await fs.open(file, 'wx');
  } catch (error) {
    if (error.code === 'EEXIST') throw new ViewerError('file_exists', 'error.fileExists', { path: file });
    throw error;
  }
  await handle.close();
  return async () => {
    const stat = await fs.stat(file).catch(() => null);
    if (stat?.isFile() && stat.size === 0) await fs.rm(file, { force: true });
  };
}

/** A free file name in `dir`, named in the language of `t`: "New deck.pptx", then "New deck (2).pptx" and so on. */
export async function suggestNewDeckPath(dir, t = tEn) {
  const folder = dir && path.isAbsolute(dir) ? dir : path.join(os.homedir(), 'Documents');
  const basename = t('newDeck.defaultName');
  for (let n = 1; n < 1000; n++) {
    const candidate = path.join(folder, `${basename}${n === 1 ? '' : ` (${n})`}.pptx`);
    if (!(await exists(candidate))) return candidate;
  }
  return path.join(folder, `${basename} ${Date.now()}.pptx`);
}

/** Checks where the new presentation goes: a .pptx (the extension is added when missing) in a folder that exists, in a file that does not. */
export async function resolveNewDeckTarget(input, baseDir) {
  const raw = String(input ?? '').trim().replace(/^"(.*)"$/, '$1').trim();
  if (!raw) throw new ViewerError('invalid_path', 'error.savePathEmpty');
  let full;
  if (path.isAbsolute(raw)) full = path.normalize(raw);
  else if (baseDir) full = path.resolve(baseDir, raw);
  else throw new ViewerError('invalid_path', 'error.savePathNotAbsolute');
  const extension = path.extname(full).toLowerCase();
  if (extension && extension !== '.pptx') throw new ViewerError('unsupported_file', 'error.saveExtension');
  if (!extension) full += '.pptx';
  const folder = await fs.stat(path.dirname(full)).catch(() => null);
  if (!folder?.isDirectory()) throw new ViewerError('folder_not_found', 'error.saveFolderMissing', { path: path.dirname(full) });
  if (await exists(full)) throw new ViewerError('file_exists', 'error.fileExists', { path: full });
  return full;
}

function describeFailure(stderr, code) {
  const text = stderr.trim();
  if (code === 3) return new ViewerError('file_exists', 'error.fileExistsShort');
  if (/0x80040154|REGDB_E_CLASSNOTREG|Class not registered/i.test(text)) {
    return new ViewerError('powerpoint_unavailable', 'error.powerpointUnavailable');
  }
  const detail = text.replace(/^HRESULT=\S+\s*/, '').split(/\r?\n/)[0].slice(0, 300);
  return detail ? new ViewerError('create_failed', 'error.createFailedDetail', { detail }) : new ViewerError('create_failed', 'error.createFailed');
}

/**
 * Creates the presentation with desktop PowerPoint and returns its path. `template` is an optional PowerPoint file or template
 * to start from; a relative one resolves against `baseDir`. An existing file is never overwritten.
 */
export async function createPresentationWithPowerPoint({ path: input, template, baseDir }) {
  if (process.platform !== 'win32') throw new ViewerError('unsupported_platform', 'error.createNeedsWindows');
  const file = await resolveNewDeckTarget(input, baseDir);
  const templateFile = template && String(template).trim() ? (await resolveDeckSource(template, baseDir)).path : '';
  const release = await reserveNewDeckFile(file);
  try {
    await runCreateScript(file, templateFile);
    if (!(await exists(file)) || (await fs.stat(file)).size === 0) throw new ViewerError('create_failed', 'error.createNoFile');
    return file;
  } catch (error) {
    await release().catch(() => {});
    throw error;
  }
}

async function runCreateScript(file, templateFile) {
  const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', SCRIPT, '-Path', file, '-Reserved', ...(templateFile ? ['-Template', templateFile] : [])];
  const child = spawn(powershellExecutable(), args, { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
  let stderr = '';
  child.stderr.setEncoding('utf8');
  child.stderr.on('data', (chunk) => {
    if (stderr.length < MAX_STDERR) stderr += chunk;
  });
  child.stdout.resume();
  let timedOut = false;
  const timer = setTimeout(() => {
    timedOut = true;
    child.kill();
  }, TIMEOUT_MS);
  const code = await new Promise((resolve, reject) => {
    child.once('error', reject);
    child.once('close', resolve);
  }).catch((error) => {
    throw new ViewerError('create_failed', 'error.powershellStartFailed', { detail: error.message });
  }).finally(() => clearTimeout(timer));
  if (timedOut) throw new ViewerError('create_failed', 'error.createTimeout');
  if (code !== 0) throw describeFailure(stderr, code);
}
