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
const {validate} = require('../scripts/publish-tarballs');

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
