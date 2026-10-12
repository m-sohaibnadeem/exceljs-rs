'use strict';

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const targets = require('./release-targets');
require('./check-artifacts')(root, pkg);
const cargo = fs.readFileSync(path.join(root, '../../Cargo.toml'), 'utf8');
assert.equal(pkg.version, /^version = "([^"]+)"/m.exec(cargo)?.[1], 'Cargo/npm version mismatch');
if (process.env.GITHUB_REF?.startsWith('refs/tags/')) {
  assert.equal(process.env.GITHUB_REF, `refs/tags/v${pkg.version}`, 'Tag/version mismatch');
}
for (const file of ['native.js', 'native.d.ts', 'README.md', 'LICENSE', 'LICENSE-EXCELJS', 'LICENSE-THIRD-PARTY', 'LICENSE-RUST-DEPENDENCIES']) {
  assert.ok(fs.statSync(path.join(root, file)).size > 0, `Missing ${file}`);
}
assert.deepEqual(Object.keys(pkg.optionalDependencies || {}).sort(),
  pkg.napi.targets.map(target => `${pkg.name}-${targets[target]}`).sort(), 'Unexpected platform dependencies');
for (const target of pkg.napi.targets) {
  const suffix = targets[target];
  assert.ok(suffix, `Unsupported target ${target}`);
  const dir = path.join(root, 'npm', suffix);
  const platform = JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8'));
  assert.equal(platform.name, `${pkg.name}-${suffix}`);
  assert.equal(platform.version, pkg.version);
  assert.equal(pkg.optionalDependencies[platform.name], pkg.version);
  assert.equal(platform.license, 'MIT');
  assert.equal(platform.main, `${pkg.napi.binaryName}.${suffix}.node`);
  assert.ok(fs.statSync(path.join(dir, platform.main)).size > 0, `Missing binary for ${target}`);
  for (const file of ['LICENSE', 'LICENSE-EXCELJS', 'LICENSE-THIRD-PARTY', 'LICENSE-RUST-DEPENDENCIES']) {
    assert.equal(fs.readFileSync(path.join(dir, file), 'utf8'), fs.readFileSync(path.join(root, file), 'utf8'));
  }
}
console.log(`Release inputs verified for ${pkg.name}@${pkg.version} (${pkg.napi.targets.length} targets).`);
