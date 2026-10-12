'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {createHash} = require('node:crypto');
const {spawnSync} = require('node:child_process');
const pkg = require('../package.json');
const targets = require('../scripts/release-targets');
const {validate, waitForPublications, publishPackages} = require('../scripts/publish-tarballs');

const examplePackages = ['linux', 'mac', 'windows', 'root'].map(name =>
  ({name, version: '1.0.0', integrity: 'sha512-' + name}));
const quiet = () => {};

test('registry waiting handles scans longer than 15 minutes and transient errors', async () => {
  let elapsed = 0;
  let calls = 0;
  const items = examplePackages.slice(0, 3);
  const available = new Set();
  await waitForPublications(items, {
    now: () => elapsed, pause: async ms => { elapsed += ms; }, log: quiet,
    lookup: async item => {
      calls++;
      if (elapsed === 0) throw Object.assign(new Error('Registry unavailable'), {retryable: true});
      if (item.name === 'linux' || elapsed >= 20 * 60000) {
        assert.ok(!available.has(item.name), 'Already verified packages should not be polled again');
        available.add(item.name);
        return item.integrity;
      }
      return null;
    },
  });
  assert.equal(elapsed, 20 * 60000);
  assert.equal(available.size, 3);
  assert.ok(calls > 3);
});

test('registry waiting rejects mismatches, fatal errors, and bounded timeouts', async () => {
  const items = [examplePackages[0]];
  await assert.rejects(waitForPublications(items, {lookup: async () => 'wrong', log: quiet}), /integrity mismatch/);
  await assert.rejects(waitForPublications(items, {lookup: async () => {throw new Error('Forbidden');}, log: quiet}), /Forbidden/);
  let elapsed = 0;
  await assert.rejects(waitForPublications(items, {
    lookup: async () => null, now: () => elapsed, pause: async ms => {elapsed += ms;}, log: quiet,
  }), /processing or review may still be pending/);
  assert.equal(elapsed, 30 * 60000);
});

test('all native uploads precede the availability gate and root publication', async () => {
  const events = [];
  await publishPackages(examplePackages, {
    lookup: async () => null, log: quiet,
    upload: async item => {events.push('upload:' + item.name);},
    wait: async items => {events.push('wait:' + items.map(item => item.name).join(','));},
  });
  assert.deepEqual(events, ['upload:linux', 'upload:mac', 'upload:windows', 'wait:linux,mac,windows', 'upload:root', 'wait:root']);
});

test('resume skips matching archives and rejects conflicts before any upload', async () => {
  const uploads = [];
  await publishPackages(examplePackages, {
    lookup: async item => item.name === 'mac' ? null : item.integrity,
    upload: async item => {uploads.push(item.name);}, wait: async () => {}, log: quiet,
  });
  assert.deepEqual(uploads, ['mac']);
  await assert.rejects(publishPackages(examplePackages, {
    lookup: async item => item.name === 'root' ? 'wrong' : null,
    upload: async () => {assert.fail('Must check every version before uploading');}, log: quiet,
  }), /differs from this archive/);
});

test('root is never uploaded when native verification fails', async () => {
  const uploads = [];
  await assert.rejects(publishPackages(examplePackages, {
    lookup: async () => null, upload: async item => {uploads.push(item.name);},
    wait: async () => {throw new Error('still processing');}, log: quiet,
  }), /still processing/);
  assert.deepEqual(uploads, ['linux', 'mac', 'windows']);
});

test('release archives require the exact package set and matching checksums', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'exceljs-rs-release-test-'));
  try {
    const packages = [...pkg.napi.targets.map(target => `${pkg.name}-${targets[target]}`), pkg.name].map(name => {
      const file = `${name}-${pkg.version}.tgz`;
      const bytes = Buffer.from(`Validation fixture for ${name}, not a publishable archive`);
      fs.writeFileSync(path.join(dir, file), bytes);
      return {name, version: pkg.version, file, integrity: 'sha512-' + createHash('sha512').update(bytes).digest('base64')};
    });
    fs.writeFileSync(path.join(dir, 'release-manifest.json'), JSON.stringify({version: pkg.version, packages}));
    assert.equal(validate(dir).length, 6);
    const verify = spawnSync(process.execPath, [path.join(__dirname, '../scripts/publish-tarballs.js'), '--verify-only'],
      {encoding: 'utf8', env: {...process.env, EXCELJS_RS_RELEASE_DIR: dir}});
    assert.equal(verify.status, 0, verify.stderr);
    fs.writeFileSync(path.join(dir, 'unexpected.tgz'), 'extra');
    assert.throws(() => validate(dir), /Unexpected archives/);
    fs.unlinkSync(path.join(dir, 'unexpected.tgz'));
    fs.appendFileSync(path.join(dir, packages[0].file), 'changed');
    assert.throws(() => validate(dir), /Archive integrity mismatch/);
    fs.unlinkSync(path.join(dir, packages[0].file));
    assert.throws(() => validate(dir), /ENOENT/);
  } finally {
    fs.rmSync(dir, {recursive: true, force: true});
  }
});
