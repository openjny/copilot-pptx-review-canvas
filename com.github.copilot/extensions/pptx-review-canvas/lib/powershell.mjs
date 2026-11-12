import path from 'node:path';

/** Windows PowerShell 5.1, which the scripts in scripts/ are written for (they are kept ASCII-only for it). */
export function powershellExecutable() {
  const systemRoot = process.env.SystemRoot || process.env.windir || 'C:\\Windows';
  return path.join(systemRoot, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
}
