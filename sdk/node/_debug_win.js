'use strict';
// TEMP Windows diagnostic — prints the pipe name the SDK computes vs the pipe the
// resident actually creates, to pinpoint a name mismatch. Delete after use.
const { spawnSync } = require('child_process');
const os = require('os');
const path = require('path');
const fs = require('fs');
const oc = require('./index.js');

const f = path.join(os.tmpdir(), `dbg-${process.pid}.xlsx`);
const [main, ping] = oc.pipePaths(f);
console.log('FILE          :', f);
console.log('SDK main pipe :', main);
console.log('SDK ping pipe :', ping);

const bin = path.join(process.env.LOCALAPPDATA || '', 'OfficeCLI', 'officecli.exe');
if (!fs.existsSync(bin)) {
  console.log('installing officecli first…');
  try { oc.install(); } catch (e) { console.log('install threw:', e.message); }
}
console.log('BIN           :', bin, 'exists:', fs.existsSync(bin));

const c = spawnSync(bin, ['create', f, '--force'], { encoding: 'utf8' });
console.log('CLI create    : status', c.status, '| stderr:', (c.stderr || '').trim().slice(0, 200));

const ps = spawnSync(
  'powershell',
  ['-NoProfile', '-Command', '(Get-ChildItem \\\\.\\pipe\\).Name | Where-Object { $_ -like "officecli*" }'],
  { encoding: 'utf8' }
);
console.log('ACTUAL pipes  :\n' + (ps.stdout || '(none)').trim());

const g = spawnSync(bin, ['get', f, '/Sheet1/A1'], { encoding: 'utf8' });
console.log('CLI get       : status', g.status, '| stdout:', (g.stdout || '').trim().slice(0, 160));

spawnSync(bin, ['close', f], { encoding: 'utf8' });
