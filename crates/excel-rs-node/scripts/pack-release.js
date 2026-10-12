'use strict';
const fs = require('node:fs');
const path = require('node:path');
const {execFileSync} = require('node:child_process');
const {createHash} = require('node:crypto');
require('./check-release');
const root = path.resolve(__dirname, '..');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const targets = require('./release-targets');
const cli = process.env.npm_execpath;
if (!cli) throw new Error('Run through npm run release:pack');
const out = path.join(root, 'release');
fs.mkdirSync(out, {recursive: true});
const packages = [];
for (const cwd of [...pkg.napi.targets.map(target => path.join(root, 'npm', targets[target])), root]) {
  const result = JSON.parse(execFileSync(process.execPath,
    [cli, 'pack', '--ignore-scripts', '--json', '--pack-destination', out], {cwd, encoding: 'utf8'}))[0];
  const bytes = fs.readFileSync(path.join(out, result.filename));
  packages.push({name: result.name, version: result.version, file: result.filename,
    integrity: 'sha512-' + createHash('sha512').update(bytes).digest('base64')});
}
fs.writeFileSync(path.join(out, 'release-manifest.json'), JSON.stringify({version: pkg.version, packages}, null, 2) + '\n');
console.log(`Packed ${packages.length} release archives with SHA-512 integrity metadata.`);
