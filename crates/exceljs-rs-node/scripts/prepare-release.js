'use strict';

const fs = require('node:fs');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const file = path.join(root, 'package.json');
const pkg = JSON.parse(fs.readFileSync(file, 'utf8'));
const targets = require('./release-targets');
// Validate the complete input set before copying files or changing manifests.
require('./check-artifacts')(root, pkg);
require('./copy-docs');
pkg.optionalDependencies = {};
for (const target of pkg.napi.targets) {
  const suffix = targets[target];
  if (!suffix) throw new Error(`Unsupported target ${target}`);
  const dir = path.join(root, 'npm', suffix);
  const platformFile = path.join(dir, 'package.json');
  const platform = JSON.parse(fs.readFileSync(platformFile, 'utf8'));
  Object.assign(platform, {
    version: pkg.version, license: pkg.license, repository: pkg.repository,
    homepage: pkg.homepage, engines: pkg.engines, publishConfig: pkg.publishConfig,
    files: [`${pkg.napi.binaryName}.${suffix}.node`, 'LICENSE', 'LICENSE-EXCELJS', 'LICENSE-THIRD-PARTY', 'LICENSE-RUST-DEPENDENCIES'],
  });
  for (const notice of ['LICENSE', 'LICENSE-EXCELJS', 'LICENSE-THIRD-PARTY', 'LICENSE-RUST-DEPENDENCIES']) {
    fs.copyFileSync(path.join(root, notice), path.join(dir, notice));
  }
  fs.writeFileSync(platformFile, JSON.stringify(platform, null, 2) + '\n');
  pkg.optionalDependencies[platform.name] = pkg.version;
}
fs.writeFileSync(file, JSON.stringify(pkg, null, 2) + '\n');
require('./check-release');
