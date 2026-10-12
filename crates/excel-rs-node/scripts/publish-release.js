'use strict';

// Publication is explicit; no npm lifecycle hook publishes other packages.
const {spawnSync} = require('node:child_process');
const path = require('node:path');
const fs = require('node:fs');
require('./check-release');
const root = path.resolve(__dirname, '..');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const targets = require('./release-targets');
const args = process.argv.slice(2);
if (args.some(arg => !['--dry-run', '--provenance'].includes(arg))) {
  throw new Error('Supported options: --dry-run, --provenance');
}
const cli = process.env.npm_execpath;
if (!cli) throw new Error('Run through npm run release:publish');
for (const cwd of [...pkg.napi.targets.map(target => path.join(root, 'npm', targets[target])), root]) {
  const result = spawnSync(process.execPath, [cli, 'publish', '--access', 'public', '--ignore-scripts', ...args],
    {cwd, stdio: 'inherit'});
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status || 1);
}
