'use strict';
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const pkg = require('../crates/excel-rs-node/package.json');
const cargo = fs.readFileSync(path.join(root, 'Cargo.toml'), 'utf8');
const lock = fs.readFileSync(path.join(root, 'Cargo.lock'), 'utf8');
assert.match(pkg.version, /^\d+\.\d+\.\d+$/, 'This workflow publishes stable versions only');
assert.equal(/^version = "([^"]+)"/m.exec(cargo)?.[1], pkg.version, 'Cargo/npm versions differ');
for (const name of ['excel-rs', 'excel-rs-node']) {
  const entry = lock.split('[[package]]').find(block => block.includes(`name = "${name}"\n`));
  assert.equal(/^version = "([^"]+)"/m.exec(entry || '')?.[1], pkg.version, `Cargo.lock version mismatch: ${name}`);
}
const npmLock = JSON.parse(fs.readFileSync(path.join(root, 'crates/excel-rs-node/package-lock.json'), 'utf8'));
assert.equal(npmLock.version, pkg.version, 'npm lockfile version mismatch');
assert.equal(npmLock.packages[''].version, pkg.version, 'npm root lockfile version mismatch');
if (process.env.GITHUB_REPOSITORY) {
  assert.equal(pkg.repository.url, `git+https://github.com/${process.env.GITHUB_REPOSITORY}.git`, 'Repository URL must match the publishing repository');
}
if (process.env.RELEASE_PUBLISH === 'true') {
  assert.equal(process.env.GITHUB_REF, `refs/tags/v${pkg.version}`, 'Publishing requires the matching version tag');
  assert.notEqual(process.env.REPOSITORY_PRIVATE, 'true', 'Publishing with provenance requires a public repository');
}
console.log(`Release configuration verified for ${pkg.name}@${pkg.version}.`);
