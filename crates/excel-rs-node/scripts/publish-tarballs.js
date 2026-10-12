'use strict';

const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const {createHash} = require('node:crypto');
const {execFileSync} = require('node:child_process');
const pkg = require('../package.json');
const targets = require('./release-targets');

function validate(dir) {
  const manifest = JSON.parse(fs.readFileSync(path.join(dir, 'release-manifest.json'), 'utf8'));
  assert.equal(manifest.version, pkg.version, 'Release version mismatch');
  if (process.env.GITHUB_REF?.startsWith('refs/tags/')) {
    assert.equal(process.env.GITHUB_REF, `refs/tags/v${pkg.version}`, 'Tag/version mismatch');
  }
  const expected = [...pkg.napi.targets.map(target => `${pkg.name}-${targets[target]}`), pkg.name];
  assert.deepEqual(manifest.packages.map(item => item.name), expected, 'Unexpected package set or publication order');
  for (const item of manifest.packages) {
    assert.equal(item.version, pkg.version);
    assert.equal(item.file, `${item.name}-${pkg.version}.tgz`);
    const actual = 'sha512-' + createHash('sha512').update(fs.readFileSync(path.join(dir, item.file))).digest('base64');
    assert.equal(actual, item.integrity, `Archive integrity mismatch: ${item.file}`);
  }
  assert.deepEqual(fs.readdirSync(dir).filter(file => file.endsWith('.tgz')).sort(),
    manifest.packages.map(item => item.file).sort(), 'Unexpected archives in release directory');
  return manifest.packages;
}

async function registryIntegrity(item) {
  const response = await fetch(`https://registry.npmjs.org/${encodeURIComponent(item.name)}/${encodeURIComponent(item.version)}`,
    {signal: AbortSignal.timeout(30000)});
  if (response.status === 404) return null;
  if (!response.ok) throw new Error(`Registry lookup failed for ${item.name}: ${response.status}`);
  const metadata = await response.json();
  return metadata.dist?.integrity || 'missing-integrity';
}

async function main() {
  const args = process.argv.slice(2);
  assert.ok(args.every(arg => ['--verify-only', '--dry-run'].includes(arg)), 'Unknown release argument');
  const dir = path.resolve(process.env.EXCEL_RS_RELEASE_DIR || path.join(__dirname, '../release'));
  const packages = validate(dir);
  if (args.includes('--verify-only')) {
    console.log(`Verified all ${packages.length} release archives.`);
    return;
  }
  const cli = process.env.npm_execpath;
  assert.ok(cli, 'Run through npm run release:publish-tarballs');
  const dryRun = args.includes('--dry-run');
  const pending = [];
  // Inspect every existing version before the first publication. A retry may
  // skip an existing package only when the tarball is byte-for-byte identical.
  for (const item of packages) {
    const integrity = dryRun ? null : await registryIntegrity(item);
    if (integrity === null) pending.push(item);
    else {
      assert.equal(integrity, item.integrity, `Published ${item.name}@${item.version} differs from this archive. Use a new version.`);
      console.log(`Already published with matching integrity: ${item.name}@${item.version}`);
    }
  }
  for (const item of pending) {
    const command = ['publish', path.join(dir, item.file), '--access', 'public', '--ignore-scripts', '--registry=https://registry.npmjs.org'];
    if (dryRun) command.push('--dry-run');
    else if (process.env.GITHUB_ACTIONS === 'true') command.push('--provenance');
    execFileSync(process.execPath, [cli, ...command], {stdio: 'inherit'});
    if (!dryRun) {
      let integrity = null;
      for (let attempt = 0; attempt < 6; attempt++) {
        integrity = await registryIntegrity(item);
        if (integrity !== null) break;
        await new Promise(resolve => setTimeout(resolve, 2000));
      }
      assert.equal(integrity, item.integrity, `Registry verification failed: ${item.name}@${item.version}`);
    }
  }
}

if (require.main === module) main().catch(error => { console.error(error.message); process.exitCode = 1; });
module.exports = {validate};
