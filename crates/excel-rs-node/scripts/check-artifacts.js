'use strict';

const fs = require('node:fs');
const path = require('node:path');
const targets = require('./release-targets');

function checkArtifacts(root, pkg) {
  const problems = [];
  const requireFile = relative => {
    const file = path.join(root, relative);
    if (!fs.existsSync(file) || !fs.statSync(file).isFile() || fs.statSync(file).size === 0) {
      problems.push(`Missing or empty: ${relative}`);
    }
  };
  for (const file of ['native.js', 'native.d.ts']) requireFile(file);
  for (const target of pkg.napi.targets) {
    const suffix = targets[target];
    if (!suffix) {
      problems.push(`Unsupported target: ${target}`);
      continue;
    }
    requireFile(`npm/${suffix}/${pkg.napi.binaryName}.${suffix}.node`);
    const relative = `npm/${suffix}/package.json`;
    try {
      const platform = JSON.parse(fs.readFileSync(path.join(root, relative), 'utf8'));
      if (platform.name !== `${pkg.name}-${suffix}`) problems.push(`Unexpected package name in ${relative}`);
    } catch {
      problems.push(`Missing or invalid: ${relative}`);
    }
  }
  if (problems.length) {
    console.error('Release blocked. Required build artifacts are unavailable:\n\n' +
      problems.map(problem => `  - ${problem}`).join('\n') +
      '\n\nA local build produces a binary for the build target only.' +
      '\nnapi artifacts collects existing binaries; it does not compile other platforms.' +
      '\nObtain tested builds for every configured target before preparing this release.' +
      '\nCreating an empty artifacts directory will not resolve missing binaries.' +
      '\nNo packages were published. See RELEASING.md at the repository root.');
    process.exit(1);
  }
}

module.exports = checkArtifacts;
