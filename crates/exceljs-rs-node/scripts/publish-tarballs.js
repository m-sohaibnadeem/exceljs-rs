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

const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));

async function registryIntegrity(item) {
  let response;
  try {
    response = await fetch(`https://registry.npmjs.org/${encodeURIComponent(item.name)}/${encodeURIComponent(item.version)}`,
      {signal: AbortSignal.timeout(30000), headers: {'Cache-Control': 'no-cache'}});
  } catch (error) {
    error.retryable = true;
    throw error;
  }
  if (response.status === 404) return null;
  if (!response.ok) {
    const error = new Error(`Registry lookup failed for ${item.name}: ${response.status}`);
    error.retryable = response.status === 429 || response.status >= 500;
    throw error;
  }
  const metadata = await response.json();
  return metadata.dist?.integrity || 'missing-integrity';
}

async function waitForPublications(items, {
  lookup = registryIntegrity, pause = sleep, now = Date.now,
  timeoutMs = 30 * 60 * 1000, intervalMs = 30000, log = console.log,
} = {}) {
  const deadline = now() + timeoutMs;
  let pending = [...items];
  while (pending.length) {
    const results = await Promise.all(pending.map(async item => {
      try {
        const integrity = await lookup(item);
        if (integrity === null) return item;
        assert.equal(integrity, item.integrity, `Registry integrity mismatch: ${item.name}@${item.version}`);
        log(`Available with matching integrity: ${item.name}@${item.version}`);
        return null;
      } catch (error) {
        if (!error.retryable) throw error;
        log(`Registry temporarily unavailable for ${item.name}; retrying.`);
        return item;
      }
    }));
    pending = results.filter(Boolean);
    if (!pending.length) return;
    if (now() >= deadline) {
      throw new Error(`npm availability timed out for: ${pending.map(item => item.name).join(', ')}. ` +
        'Uploads were accepted, but processing or review may still be pending. Check npm package status, then resume with the same artifacts.');
    }
    log(`Waiting for npm processing: ${pending.map(item => item.name).join(', ')}`);
    await pause(Math.min(intervalMs, deadline - now()));
  }
}

async function publishPackages(packages, {lookup = registryIntegrity, upload,
  wait = waitForPublications, pause = sleep, log = console.log} = {}) {
  const pending = new Set();
  // Check all versions before any upload. Never accept different bytes.
  for (const item of packages) {
    let integrity;
    for (let attempt = 0; ; attempt++) {
      try { integrity = await lookup(item); break; }
      catch (error) {
        if (!error.retryable || attempt === 3) throw error;
        await pause(5000);
      }
    }
    if (integrity === null) pending.add(item);
    else {
      assert.equal(integrity, item.integrity, `Published ${item.name}@${item.version} differs from this archive.`);
      log(`Already published with matching integrity: ${item.name}@${item.version}`);
    }
  }
  const platforms = packages.slice(0, -1);
  const root = packages.at(-1);
  // Submit every platform before waiting, so npm can process them together.
  // Failed upload commands are not blindly retried: acceptance may be ambiguous.
  for (const item of platforms) if (pending.has(item)) await upload(item);
  await wait(platforms, {lookup, pause, log});
  if (pending.has(root)) await upload(root);
  await wait([root], {lookup, pause, log});
}

async function main() {
  const args = process.argv.slice(2);
  assert.ok(args.every(arg => ['--verify-only', '--dry-run'].includes(arg)), 'Unknown release argument');
  const dir = path.resolve(process.env.EXCELJS_RS_RELEASE_DIR || path.join(__dirname, '../release'));
  const packages = validate(dir);
  if (args.includes('--verify-only')) {
    console.log(`Verified all ${packages.length} release archives.`);
    return;
  }
  const cli = process.env.npm_execpath;
  assert.ok(cli, 'Run through npm run release:publish-tarballs');
  const dryRun = args.includes('--dry-run');
  const upload = item => {
    const command = ['publish', path.join(dir, item.file), '--access', 'public', '--ignore-scripts', '--registry=https://registry.npmjs.org'];
    if (dryRun) command.push('--dry-run');
    else if (process.env.GITHUB_ACTIONS === 'true') command.push('--provenance');
    execFileSync(process.execPath, [cli, ...command], {stdio: 'inherit'});
  };
  if (dryRun) {
    for (const item of packages) upload(item);
  } else {
    await publishPackages(packages, {upload});
  }
}

if (require.main === module) main().catch(error => { console.error(error.message); process.exitCode = 1; });
module.exports = {validate, waitForPublications, publishPackages};
